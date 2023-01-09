# CAM generate Setup sheets

# For this sample script to run, the active Fusion document must contain at least one CAM operation.

import adsk.core, adsk.fusion, adsk.cam, traceback, os

############################################################ PARAMETERS ############################################################
# Roughing Sequence
Spindle_Speed_1 = "4000 rpm"
Ramp_Spindle_Speed_1 = "4000 rpm"
Lead_In_Feed_Rate_1 = "10 in/min"
Lead_Out_Feed_Rate_1 = '10in/min'
Ramp_Feed_Rate_1 = "40in/min"
Clearance_Height_1 = "0.2in"
Optimal_Load_1 = "0.12in"
Max_Step_Down_1 = "0.12in"

# Re-roughing Sequence
Spindle_Speed_2 = "4000 rpm"
Ramp_Spindle_Speed_2 = "4000 rpm"
Cutting_Feedrate_2 = "9in/min"
Lead_In_Feed_Rate_2 = "9 in/min"
Lead_Out_Feed_Rate_2 = '9in/min'
Ramp_Feed_Rate_2 = "40in/min"
Clearance_Height_2 = "0.2in"
Optimal_Load_2 = "0.05in"
Max_Step_Down_2 = "0.05in"

# Finishing Sequence
Spindle_Speed_3 = "4000 rpm"
Ramp_Spindle_Speed_3 = "4000 rpm"
Cutting_Feedrate_3 = "15in/min"
Lead_In_Feed_Rate_3 = "15 in/min"
Lead_Out_Feed_Rate_3 = '15in/min'
Max_Step_Down_3 = "0.05in"
Max_Step_Over_3 = "0.05in"



#####################################################################################################################################


# Check if parameter enabled using isEnabled

# Change doc to imperial units



def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        doc = app.activeDocument
        products = doc.products
        
        product = products.itemByProductType('CAMProductType')

        # check if the document has a CAMProductType.  It will not if there are no CAM operations in it.
        if not product:
            ui.messageBox('There are no CAM operations in the active document.  This script requires the active document to contain at least one CAM operation.',
                            'No CAM Operations Exist',
                            adsk.core.MessageBoxButtonTypes.OKButtonType,
                            adsk.core.MessageBoxIconTypes.CriticalIconType)
            return

        cam = adsk.cam.CAM.cast(product)


        # List of all setups
        setups = cam.setups
        for setup in setups:
            
            for operation in setup.operations:

                def check_vals(operation_name, target):
                    Param = operation.parameters.itemByName(operation_name)
                    x1 = target.replace(" ","")
                    x2 = Param.expression.replace(" ","")

                    if x1 != x2:
                        ui.messageBox("Error detected at {}, \n Target Value: {} \n Current Value: {}".format(Param.title,x1,x2))
                    else:
                        ui.messageBox("Same detected at {}, \n Target Value: {} \n Current Value: {}".format(Param.title,x1,x2)) 


                operation_name = operation.name
                ui.messageBox(operation_name)
                
                check_vals('tool_rampSpindleSpeed',Ramp_Spindle_Speed_1)
                check_vals('tool_feedEntry',Lead_In_Feed_Rate_1)
                check_vals('tool_feedExit',Lead_Out_Feed_Rate_1)

                if operation.parameters.itemByName("useStockToLeave").expression == 'false':
                    ui.messageBox("Stock to leave is False")
                else:
                    ui.messageBox("stock is True")
                



                # toleranceParam = operation.parameters.itemByName('tool_feedEntry')
                # if toleranceParam.expression != Lead_In_Feed_Rate:
                #     ui.messageBox("Error detected at Lead_In_Feed_Rate, \n Target Value: {} \n Current Value: {}".format(Lead_In_Feed_Rate, toleranceParam.expression )) #, ". Got:", toleranceParam.expression)
                #     return
                # toleranceParam = operation.parameters.itemByName('tolerance')
                # toleranceParam.expression = "0.336mm"

     

        # specify the output folder and format for the setup sheets
        outputFolder = cam.temporaryFolder
        sheetFormat = adsk.cam.SetupSheetFormats.HTMLFormat

        #sheetFormat = adsk.cam.SetupSheetFormats.ExcelFormat (not currently supported on Mac)

        # prompt the user with an option to view the resulting setup sheets.
        viewResults = ui.messageBox('View setup sheets when done?', 'Generate Setup Sheets',
                                    adsk.core.MessageBoxButtonTypes.YesNoButtonType,
                                    adsk.core.MessageBoxIconTypes.QuestionIconType)
        if viewResults == adsk.core.DialogResults.DialogNo:
            viewResult = False
        else:
            viewResult = True

        # set the value of scenario to 1, 2 or 3 to generate setup sheets for all, for the first setup, or for the first operation of the first setup.
        scenario = 1
        if scenario == 1:
            ui.messageBox('Setup sheets for all operations will be generated.')
            cam.generateAllSetupSheets(sheetFormat, outputFolder, viewResult)
        elif scenario == 2:
            ui.messageBox('Setup sheets for operations in the first setup will be generated.')
            setup = cam.setups.item(0)
            cam.generateSetupSheet(setup, sheetFormat, outputFolder, viewResult)
        elif scenario == 3:
            ui.messageBox('A setup sheet for the first operation in the first setup will be generated.')
            setup = cam.setups.item(0)
            operations = setup.allOperations
            operation = operations.item(0)
            if operation.hasToolpath:
                cam.generateSetupSheet(operation, sheetFormat, outputFolder, viewResult)
            else:
                ui.messageBox('This operation has no toolpath.  A valid toolpath must exist in order for a setup sheet to be generated.')
                return


    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
