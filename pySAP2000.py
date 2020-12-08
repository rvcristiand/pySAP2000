import xlwings as xw
import comtypes.client


def get_active_sap2000():
    # create API helper object
    helper = comtypes.client.CreateObject('SAP2000v1.Helper')
    helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

    # attach to a running instance of SAP2000
    try:
        # get the active SapObject
        return helper.GetObject("CSI.SAP2000.API.SapObject") 
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")

mySapObject = get_active_sap2000()

# create SapModel object
sapModel = mySapObject.SapModel

# deselect all cases and combos
ret = sapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()

# get load cases names
# ret = sapModel.LoadCases.GetNameList_1()
casos_carga = ['D', 'Lr', 'G', 'W', 'Fx', 'Fy']

# set load cases for output
for caso_carga in casos_carga:
    ret = sapModel.Results.Setup.SetCaseSelectedForOutput(caso_carga)
apoyos = [41, 51, 61, 71, 46, 56, 66, 76]

ret = sapModel.Results.JointReact('apoyos', 2)

no_resultados = ret[0]
resultados = [ret[1],
              ret[3],
              ret[6],
              ret[7],
              ret[8],
              ret[9],
              ret[10],
              ret[11]]


# create report
wb = xw.Book()
sht = wb.sheets['Sheet1']

sht.range('A1').value = ['nodo',
                         'caso de carga',
                         'fx',
                         'fy',
                         'fz',
                         'mx',
                         'my',
                         'mz']

for i in range(len(resultados)):
    sht.range('A2').offset(0, i).options(transpose=True).value = resultados[i]
