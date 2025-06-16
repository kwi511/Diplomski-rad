import pandas as pd
import math

# hourly wages

hourly_wage_RRT_SysAdmin = 39.59
hourly_wage_Org_SysAdmin = 39.59
hourly_wage_RRT_ReverseEngineer = 54.28
hourly_wage_RRT_LogAnalyst = 20.49
hourly_wage_RRT_Forensic_Computer_Analyst = 36.25
hourly_wage_employee = 25.36


file = pd.read_excel('path/to/file.xlsx', sheet_name='Chronology')


stupac_type = file.columns[0]
stupac_beginning = file.columns[1]
stupac_name = file.columns[3]
stupac_ending = file.columns[4]
stupac_actor = file.columns[6]
stupac_player = file.columns[7]


# ------------------------------------------------------------------
# 1) RRT work time

#SysAdmin

worktime_RRT_SysAdmin = 0

for index, row in file.iterrows():
    if row[stupac_type] == 'Action' and row[stupac_name] != 'Relocate Actor' and (row[stupac_actor] == 'RRT_SystemAdmin1' or row[stupac_actor] == 'RRT_SystemAdmin2') and row[stupac_player] == 'IT':
   
        beginning_dt = pd.to_datetime(row[stupac_beginning], dayfirst=True)
        ending_dt = pd.to_datetime(row[stupac_ending], dayfirst=True)
        
        trajanje_sati = (ending_dt - beginning_dt).total_seconds() / 3600
        worktime_RRT_SysAdmin += trajanje_sati


cost_RRT_SysAdmin = math.ceil(worktime_RRT_SysAdmin) * hourly_wage_RRT_SysAdmin


#ReverseEngineer

worktime_RRT_ReverseEngineer = 0

for index, row in file.iterrows():
    if row[stupac_type] == 'Action' and row[stupac_name] != 'Relocate Actor' and (row[stupac_actor] == 'RRT_Reversing1' or row[stupac_actor] == 'RRT_Reversing2') and row[stupac_player] == 'IT':
   
        beginning_dt = pd.to_datetime(row[stupac_beginning], dayfirst=True)
        ending_dt = pd.to_datetime(row[stupac_ending], dayfirst=True)

        
   
        trajanje_sati = (ending_dt - beginning_dt).total_seconds() / 3600
        worktime_RRT_ReverseEngineer += trajanje_sati


cost_RRT_ReverseEngineer = math.ceil(worktime_RRT_ReverseEngineer) * hourly_wage_RRT_ReverseEngineer



#LogAnalyst

worktime_RRT_LogAnalyst = 0

for index, row in file.iterrows():
    if row[stupac_type] == 'Action' and row[stupac_name] != 'Relocate Actor' and (row[stupac_actor] == 'RRT_LogAnalysis1' or row[stupac_actor] == 'RRT_LogAnalysis2') and row[stupac_player] == 'IT':
   
        beginning_dt = pd.to_datetime(row[stupac_beginning], dayfirst=True)
        ending_dt = pd.to_datetime(row[stupac_ending], dayfirst=True)

        trajanje_sati = (ending_dt - beginning_dt).total_seconds() / 3600
        worktime_RRT_LogAnalyst += trajanje_sati

cost_RRT_LogAnalyst = math.ceil(worktime_RRT_LogAnalyst) * hourly_wage_RRT_LogAnalyst



#ForensicComputerAnalyst

worktime_RRT_ForensicComputerAnalyst = 0

for index, row in file.iterrows():
    if row[stupac_type] == 'Action' and row[stupac_name] != 'Relocate Actor' and (row[stupac_actor] == 'RRT_Forensics1' or row[stupac_actor] == 'RRT_Forensics2') and row[stupac_player] == 'IT':
   
        beginning_dt = pd.to_datetime(row[stupac_beginning], dayfirst=True)
        ending_dt = pd.to_datetime(row[stupac_ending], dayfirst=True)


        trajanje_sati = (ending_dt - beginning_dt).total_seconds() / 3600
        worktime_RRT_ForensicComputerAnalyst += trajanje_sati

cost_RRT_ForensicComputerAnalyst = math.ceil(worktime_RRT_ForensicComputerAnalyst) * hourly_wage_RRT_Forensic_Computer_Analyst


#total cost

total_cost_RRT = cost_RRT_SysAdmin + cost_RRT_ReverseEngineer + cost_RRT_LogAnalyst + cost_RRT_ForensicComputerAnalyst
print('total_cost_RRT: {}'.format(total_cost_RRT))

#-------------------------------------------------------------

# 2) employee work time

worktime_sysadmin_organization = 0
worktime_organization = 0

for index, row in file.iterrows():
    if row[stupac_type] == 'Action' and row[stupac_name] != 'Relocate Actor' and row[stupac_name] != 'Request Assistance' and row[stupac_player] == 'IT':
        if pd.notna(row[stupac_actor]):
            actor_value = row[stupac_actor].lower()
            if 'attacker' not in actor_value and 'rrt' not in actor_value:
                if 'n&sa' in actor_value:
                    beginning_dt = pd.to_datetime(row[stupac_beginning], dayfirst=True)
                    ending_dt = pd.to_datetime(row[stupac_ending], dayfirst=True)


                    trajanje_sati = (ending_dt - beginning_dt).total_seconds() / 3600
                    worktime_sysadmin_organization += trajanje_sati
                else:
                    beginning_dt = pd.to_datetime(row[stupac_beginning], dayfirst=True)
                    ending_dt = pd.to_datetime(row[stupac_ending], dayfirst=True)


                    trajanje_sati = (ending_dt - beginning_dt).total_seconds() / 3600
                    worktime_organization += trajanje_sati
                    


total_cost_organization = (math.ceil(worktime_sysadmin_organization) * hourly_wage_Org_SysAdmin)  + (math.ceil(worktime_organization) * hourly_wage_employee)
print('total_cost_employee_time: {}'.format(total_cost_organization))



