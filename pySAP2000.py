import comtypes.client


# create API helper object
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

# attach to a running instance of SAP2000
try:
    # get the active SapObject
    mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject") 
except (OSError, comtypes.COMError):
    print("No running instance of the program found or failed to attach.")

# create SapModel object
sapModel = mySapObject.SapModel

# get load patterns
# _, load_patterns, ret = sapModel.LoadPatterns.GetNameList()

# get load cases
# _, load_cases, ret = sapModel.LoadCases.GetNameList_1()
# print(load_cases)    

# create load combinations
# ---------------------

# get load combinations
load_combinations = ()
try:
    _, load_combinations, ret = sapModel.RespCombo.GetNameList()
except IndexError:
    print("no load cases defined")

# delete ALL load cases
if load_combinations:    
    for combination in load_combinations:
        ret = sapModel.RespCombo.Delete(combination)

# create load combinations
combinations = {
    # esfuerzos admisibles
    "D": [(1.0, 'D')],

    "D+Lr": [(1.0, 'D'), (1.0, 'Lr')],
    "D+G": [(1.0, 'D'), (1.0, 'G')],

    "D+W": [(1.0, 'D'), (1.0, 'W')],

    "D+Fx+0.3Fy": [(1.0, 'D'), (1.0, 'Fx'), (0.3, 'Fy')],
    "D+0.3Fx+Fy": [(1.0, 'D'), (0.3, 'Fx'), (1.0, 'Fy')],

    "D+0.75W+0.75Lr": [(1.0, 'D'), (0.75, 'W'), (0.75, 'Lr')],
    "D+0.75W+0.75G": [(1.0, 'D'), (0.75, 'W'), (0.75, 'G')],

    "D+0.75Fx+0.225Fy+0.75Lr": [(1.0, 'D'), (0.75, 'Fx'), (0.225, 'Fy'), (0.75, 'Lr')],
    "D+0.75Fx+0.225Fy+0.75G": [(1.0, 'D'), (0.75, 'Fx'), (0.225, 'Fy'), (0.75, 'G')],
    "D+0.225Fx+0.75Fy+0.75Lr": [(1.0, 'D'), (0.225, 'Fx'), (0.75, 'Fy'), (0.75, 'Lr')],
    "D+0.225Fx+0.75Fy+0.75G": [(1.0, 'D'), (0.225, 'Fx'), (0.75, 'Fy'), (0.75, 'G')],

    "0.6D-1.0W": [(0.6, 'D'), (-1.0, 'W')],

    "1.4D": [(1.4, 'D')],
    
    "1.2D+0.5Lr": [(1.2, 'D'), (0.5, 'Lr')],
    "1.2D+0.5G": [(1.2, 'D'), (0.5, 'G')],
    
    "1.2D+1.6Lr+0.5W": [(1.2, 'D'), (1.6, 'Lr'), (0.5, 'W')],
    "1.2D+1.6G+0.5W": [(1.2, 'D'), (1.6, 'G'), (0.5, 'W')],

    "1.2D+1.0W+0.5Lr": [(1.2, 'D'), (1.0, 'W'), (0.5, 'Lr')],
    "1.2D+1.0W+0.5G": [(1.2, 'D'), (1.0, 'W'), (0.5, 'G')],

    "1.2D+1.0Ex+0.3Ey": [(1.2, 'D'), (1.0, 'Ex'), (0.3, 'Ey')],
    "1.2D+1.0Ex+0.3Ey": [(1.2, 'D'), (1.0, 'Ex'), (0.3, 'Ey')],

    "0.9D-1.0W": [(0.9, 'D'), (-1.0, 'W')]
}

for key, value in combinations.items():
    sapModel.RespCombo.Add(key, 0)
    for load in value:
        sapModel.RespCombo.SetCaseList(key, 0, load[1], load[0])