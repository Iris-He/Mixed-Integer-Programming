from pyomo.core import *
from pyomo.opt import SolverFactory
from pyomo.environ import *
import pandas as pd

#IMPORTANT: Data of the same network should be continous in the input Excel Table

def preprocess_Data(Data,Networks_list):
    # This preprocess was added because if not the read of Data for too time consuming since 
    # it had to revise all position in the Dataframe several times
    
    dict = {}
    for n in Networks_list:
        pos = [0,0]
        i = 0
        initial = False
        final = False
        while i < len(Data):
            if initial == False:
                if Data['NETWORK_ID'].at[i] == n:
                    pos[0] = i
                    initial = True
            else:
                if Data['NETWORK_ID'].at[i] != n:
                    pos[1] = i
                    dict[n] = pos
                    break
            if i == len(Data) - 1: # For the last network in the Networks_list
                pos[1] = i + 1
                dict[n] = pos
                break
            i += 1
    return dict

def run_store_allocation(file,sheet,relaxation_Rank):

    #Read of data and save of the list of stores and networks as python lists
    Data = pd.read_excel(file, sheet_name=sheet, engine='openpyxl')
    Networks_list = Data['NETWORK_ID'].tolist()
    Networks_list = set(Networks_list)
    Networks_list = list(Networks_list)
    Stores_list = Data['STORE_ID'].tolist()
    Stores_list = set(Stores_list)
    Stores_list = list(Stores_list)
 
    # Creation of the Concrete Model
    model = ConcreteModel()

    # Creation of sets
    model.Networks = Set(initialize=Networks_list, doc='Stores')
    model.Stores = Set(initialize=Stores_list, doc='Networks')

    # Call the preprocess
    dict = preprocess_Data(Data,Networks_list)

    # Naive rank. If it appears in the solution it is because the real problem was infeasible due to the
    # limit in the number of times that a store can be picked as a hub
    naive_Rank = relaxation_Rank * 10000

    # Build of parameter Rank
    def rule_Rank(model,s,n):
        limits = dict[n]
        for i in range(limits[0],limits[1]):
            if Data['NETWORK_ID'].at[i] == n and Data['STORE_ID'].at[i] == s:
                return int(Data['RANK'].at[i])
        else:
            return naive_Rank
    model.Rank = Param(model.Stores, model.Networks, initialize=rule_Rank, doc='Ranks')

    # Build of parameter of the number of hubs needed by each Network
    def rule_StoresNeed(model,n):
        limits = dict[n]
        for i in range(limits[0],limits[1]):
            if Data['NETWORK_ID'].at[i] == n:
                return int(Data['hubs needed'].at[i])

    model.StoresNeed = Param(model.Networks, initialize=rule_StoresNeed, doc='Needs')
    
    # Creation of assigment variable. Since in this python program it is formulated as an assgiment problem
    # there is no need of specifying a binary nature. It is enough to work with a continous variable
    # which have the bounds 0 and 1 due to the mathematical properties of the restrictions. This is faster to solve
    # than addressing an NP-hard integer programming model and the results are the same.
    model.x = Var(model.Stores, model.Networks, domain=NonNegativeReals, bounds=(0,1), doc='Assignment')

    # Relaxation variable since with the limit of three times a store can be a hub there was no feasible solution
    # according to the provided input data. This decision was made after a short talk with Iris.
    model.extra = Var(model.Stores, domain=NonNegativeReals, doc='Relaxation')

    # Constraint for the maximum number of times that a store can be selected as hub, including a relaxation term.
    def maximum_assign_rule(model, s):
        return sum(model.x[s,n] for n in model.Networks) <= 3 + model.extra[s]
    model.maximum_assign = Constraint(model.Stores, rule=maximum_assign_rule, doc='Available production time at Departments')

    # Constraint of number of hubs needed by each Network
    def need_assign_rule(model, n):
        return sum(model.x[s,n] for s in model.Stores) >= model.StoresNeed[n]
    model.need_assign = Constraint(model.Networks, rule=need_assign_rule, doc='Available production time at Departments')

    # Objective function which minimizes the overall some of ranks in order to aim at selecting s hubs those stores
    # that are best ranked. It also includes the relaxation variable with a cost that it's much higher than any rank
    # so it works as a relaxation and a store is picked as a hub more than three times only when it is inevitable
    def objective_rule(model):
        return sum((model.x[s,n] * model.Rank[s,n] + model.extra[s] * relaxation_Rank) for s in model.Stores for n in model.Networks)
    model.objective = Objective(rule=objective_rule, sense=minimize, doc='Define objective function')

    # Solve the problem using CBC, which is an open source solver
    opt = SolverFactory("cbc")
    results = opt.solve(model)

    # Write some basic data of the resolution process on the terminal
    results.write()
    
    # Print the results in an organized way in a results.txt file
    with open("results.txt", "w") as f:
        for n in model.Networks:
            f.write("Network %s has the following stores as hubs: \n" % str(n))
            for s in model.Stores:
                if model.x[s,n].value > 0.5:
                    f.write("Store %s is ranked with %i in this Network (%.2f) \n" % (str(s),int(model.Rank[s,n]),model.x[s,n].value))
        for s in model.Stores:   
            if model.extra[s].value > 0.5: 
                f.write("Needed extra of %.2f in store %s: \n" % (model.extra[s].value,str(s)))

def main():

    # Set the Excel file and the sheet where is the Data. Remember to filter beforehand networks which require
    # more hubs than available stores. In case this problem is not correct
    # you can detected in a lately stage when you see that the solver pick assigments with extralarger naive cost.
    file = "Input_Data.xlsx"
    sheet = "Data_filtered"

    # Pick a cost of relaxation. The model includes a relaxation variable with a rank that it's much higher than any rank
    # so it works as a relaxation and a store is picked as a hub more than three times only when it is inevitable
    relaxation_Rank = 1000

    run_store_allocation(file,sheet,relaxation_Rank)


if __name__ == '__main__':

    main()