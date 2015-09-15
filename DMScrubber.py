import pandas
pandas.options.mode.chained_assignment = None

class CoreRow:
	def __init__(self,excel_input):
		self.StowStrategy = excel_input['Stow Strategy'].iloc[0]
		self.NumberofModules = excel_input['Number of Modules'].iloc[0]
		self.SnowLoad = excel_input['Snow Load'].iloc[0]
		self.PastProjects = excel_input['Past Projects'].iloc[0]
		self.RowInformation = excel_input['Row Information'].iloc[0]
		self.InteriororExterior = excel_input['Interior or Exterior'].iloc[0]
		self.Mksi = []
		self.Bay1_length = excel_input['Bay 1'].iloc[0]
		self.Bay1_ksi = excel_input['Bay 1'].iloc[1]
		self.Bay2_length = excel_input['Bay 2'].iloc[0]
		self.Bay2_ksi = excel_input['Bay 2'].iloc[1]
		self.Bay3_length = excel_input['Bay 3'].iloc[0]
		self.Bay3_ksi = excel_input['Bay 3'].iloc[1]
		self.Bay4_length = excel_input['Bay 4'].iloc[0]
		self.Bay4_ksi = excel_input['Bay 4'].iloc[1]
		self.Bay5_length = excel_input['Bay 5'].iloc[0]
		self.Bay5_ksi = excel_input['Bay 5'].iloc[1]
		self.Bay6_length = excel_input['Bay 6'].iloc[0]
		self.Bay6_ksi = excel_input['Bay 6'].iloc[1]
		self.Bay7_length = excel_input['Bay 7'].iloc[0]
		self.Bay7_ksi = excel_input['Bay 7'].iloc[1]
		self.Bay8_length = excel_input['Bay 8'].iloc[0]
		self.Bay8_ksi = excel_input['Bay 8'].iloc[1]
		self.Bay9_length = excel_input['Bay 9'].iloc[0]
		self.Bay9_ksi = excel_input['Bay 9'].iloc[1]
		self.Bay10_length = excel_input['Bay 10'].iloc[0]
		self.Bay10_ksi = excel_input['Bay 10'].iloc[1]
		self.Maxksi = excel_input['Max ksi'].iloc[1]
		self.ExistingCoreBlocks = excel_input['Existing Core Blocks'].iloc[0]
		self.EngineeringCheck = excel_input['Engineering Check'].iloc[0]
		self.EngineeringNotes = excel_input['Engineering Notes'].iloc[0]
		self.OperationsCheck = excel_input['Operations Check'].iloc[0]
		self.OperationsNotes = excel_input['Operations Notes'].iloc[0]
		self.SalesCheck = excel_input['Sales Check'].iloc[0]
		self.SalesNotes = excel_input['Sales Notes'].iloc[0]
		self.EngNumberofModules = excel_input['Number of Modules'].iloc[0]
		self.EngInteriororExterior = excel_input['Eng Interior or Exterior/Tube Thickness'].iloc[0]
		self.EngTubeThickness = excel_input['Eng Interior or Exterior/Tube Thickness'].iloc[1]
		self.EngArrayMotor = excel_input['Eng Array/Motor'].iloc[0]
		self.EngPierType = excel_input['Eng Array/Motor'].iloc[1]
		self.EngMksi = []
		self.EngBay1_length = excel_input['Eng Bay 1'].iloc[0]
		self.EngBay1_ksi = excel_input['Eng Bay 1'].iloc[1]
		self.EngBay2_length = excel_input['Eng Bay 2'].iloc[0]
		self.EngBay2_ksi = excel_input['Eng Bay 2'].iloc[1]
		self.EngBay3_length = excel_input['Eng Bay 3'].iloc[0]
		self.EngBay3_ksi = excel_input['Eng Bay 3'].iloc[1]
		self.EngBay4_length = excel_input['Eng Bay 4'].iloc[0]
		self.EngBay4_ksi = excel_input['Eng Bay 4'].iloc[1]
		self.EngBay5_length = excel_input['Eng Bay 5'].iloc[0]
		self.EngBay5_ksi = excel_input['Eng Bay 5'].iloc[1]
		self.EngBay6_length = excel_input['Eng Bay 6'].iloc[0]
		self.EngBay6_ksi = excel_input['Eng Bay 6'].iloc[1]
		self.EngBay7_length = excel_input['Eng Bay 7'].iloc[0]
		self.EngBay7_ksi = excel_input['Eng Bay 7'].iloc[1]
		self.EngBay8_length = excel_input['Eng Bay 8'].iloc[0]
		self.EngBay8_ksi = excel_input['Eng Bay 8'].iloc[1]
		self.EngBay9_length = excel_input['Eng Bay 9'].iloc[0]
		self.EngBay9_ksi = excel_input['Eng Bay 9'].iloc[1]
		self.EngBay10_length = excel_input['Eng Bay 10'].iloc[0]
		self.EngBay10_ksi = excel_input['Eng Bay 10'].iloc[1]
		self.EngMaxksi = excel_input['Eng Max ksi'].iloc[1]
		self.EngDirectFasten = excel_input['Eng Direct Fasten'].iloc[0]
		self.EngPanelRail = excel_input['Eng Panel Rail'].iloc[0]
		self.EngUBoltSA = excel_input['Eng U-Bolt SA'].iloc[0]
		self.EngTorqueTube = excel_input['Eng Torque Tube'].iloc[0]
		self.EngTTBolts = excel_input['Eng TT Bolts'].iloc[0]
		self.EngCastRail = excel_input['Eng Cast Rail'].iloc[0]
		self.EngPivotPin = excel_input['Eng Pivot Pin'].iloc[0]
		self.EngHandle = excel_input['Eng Handle'].iloc[0]
		self.EngTTTADrive = excel_input['Eng TTA Drive'].iloc[0]
		self.EngTTADABolts = excel_input['Eng TTADA Bolts'].iloc[0]
		self.EngMotorMount = excel_input['Eng Motor Mount'].iloc[0]
		self.EngArrayPier = excel_input['Eng Array Pier'].iloc[0]
		self.EngMotorPier = excel_input['Eng Motor Pier'].iloc[0]

DMpath = "C:\Users\knotohamiprodjo\Desktop\py_data"
DMfilename = "\NEXTracker Design Matrix TESTER.xlsx"

excel = pandas.read_excel(DMpath + DMfilename, sheetname = "100mph",skiprows = [0,1])

column_names = [
	'Stow Strategy',#A
	'Number of Modules',#B
	'Number of Piers',#C
	'Snow Load',#D
	'Past Projects',#E
	'Row Information',#F
	'Interior or Exterior',#G
	'M/ksi',#H
	'Bay 1',#I
	'Bay 2',#J
	'Bay 3',#K
	'Bay 4',#L
	'Bay 5',#M
	'Bay 6',#N
	'Bay 7',#O
	'Bay 8',#P
	'Bay 9',#Q
	'Bay 10',#R
	'Max ksi',#S
	'Existing Core Blocks',#T
	'Engineering Check',#U
	'Engineering Notes',#V
	'Operations Check',#W
	'Operations Notes',#X
	'Sales Check',#Y
	'Sales Notes',#Z
	'Eng Number of Modules',#AA
	'Eng Interior or Exterior/Tube Thickness',#AB
	'Eng Array/Motor',#AC
	'Eng Pier Type',#AD
	'Eng M/ksi',#AE
	'Eng Bay 1',#AF
	'Eng Bay 2',#AG
	'Eng Bay 3',#AH
	'Eng Bay 4',#AI
	'Eng Bay 5',#AJ
	'Eng Bay 6',#AK
	'Eng Bay 7',#AL
	'Eng Bay 8',#AM
	'Eng Bay 9',#AN
	'Eng Bay 10',#AO
	'Eng Max ksi',#AP
	'Eng Direct Fasten',#AQ
	'Eng Panel Rail',#AR
	'Eng U-Bolt SA',#AS
	'Eng Torque Tube',#AT
	'Eng TT Bolts',#AU
	'Eng Cast Rail',#AV
	'Eng Pivot Pin',#AW
	'Eng Handle',#AX
	'Eng TTA Drive',#AY
	'Eng TTADA Bolts',#AZ
	'Eng Motor Mount',#BA
	'Eng Array Pier',#BB
	'Eng Motor Pier'#BC
	]

excel.columns = column_names
padded_excel = excel.fillna(method = 'pad')

excel.update(padded_excel[[column_names[0],column_names[1],column_names[2]]])
separated_designs = []
for i in range(0,len(excel['Stow Strategy'])/4):
	separated_designs.append([excel[i*4:i*4+2],excel[i*4+2:i*4+4]])
for j in range(0,len(separated_designs)):
	snow_load = separated_designs[j][1].loc[:,('Snow Load')].iloc[0]
	if snow_load == '-':
		snow_load_return = 0
	else:
		if len(snow_load)>3:
			snow_load_return = snow_load[1:3]
		elif len(snow_load) == 3:
			snow_load_return = snow_load[1:]
		else:
			snow_load_return = snow_load.pop('psf')
	separated_designs[j][0].loc[:,('Snow Load')].iloc[0] = snow_load_return
	separated_designs[j][1].loc[:,('Snow Load')].iloc[0] = snow_load_return

classed_rows = []
for i in separated_designs:
	for j in i:
		classed_rows.append(CoreRow(j))

a = classed_rows[0]

# for i in cols_to_pad:
# 	excel[column_names_dict[i]] = 