
import smartsheet

#create the client with token - be sure to safeguard the token:
smartsheet_client = smartsheet.Smartsheet('kt0exv5ks4jy5c0aymg6nxwnxp')

# Get current user
user_profile = smartsheet_client.Users.get_current_user()
#list all the workspaces (this is useful to get the id of the workspace)
x = smartsheet_client.Workspaces.list_workspaces(include_all=True)
#list all the worksheets (also sometimes useful to get the ids of the sheets - but there are a lot of them!)
y = response = smartsheet_client.Sheets.list_sheets()

#create a new workspace
workspace = smartsheet_client.Workspaces.create_workspace(
  smartsheet.models.Workspace({
    'name': 'Nurse Turnover'
  })
)

#define a variable with the sheet name
sheet_name = "new test sheet"

#define the specs for a new sheet you want to create
sheet_spec = smartsheet.models.Sheet({
  'name': sheet_name,
  'columns': [{
      'title': 'Unit Name',
      'type': 'TEXT_NUMBER'
    }, {
      'title': 'Person Name',
      'primary': True,
      'type': 'TEXT_NUMBER'
    }, {
      "title": "Reason for Turnover",
      "type": "MULTI_PICKLIST",
      "options": [
        "Left Hospital",
        "left Unit",
        "left direct patient care",
        "won the lottery",
        "fired",
        "gave up; nursing too difficult"
      ],
      "width": 150
    }
  ]
})

#create the response for creating new sheet that you provided specs for in the workspace defined by the id number
response = smartsheet_client.Workspaces.create_sheet_in_workspace(
    2286958518003588,           # workspace_id
    sheet_spec)
#create the new sheet
new_sheet = response.result

sheet_id = new_sheet.id

# get primary column id
for idx, col in enumerate(new_sheet.columns):
    #print(idx)
    print(col)
    if col.primary:
        break
    sheet_primary_col = col

action = new_sheet.add_rows([smartsheet.models.Row({
'to_top': True,
'cells': [{
'column_id': sheet_primary_col.id,
'value': 'The first column of the first row.'
}]
})])

#search for that sheet by name
search_results = smartsheet_client.Search.search(sheet_name).results
#define "sheet_id as an object that is an oject_id and has the object type of a "sheet"
sheet_id = next(result.object_id for result in search_results if result.object_type == 'sheet')
#define the my_new_sheet as all the objects with an id in the sheet
my_new_sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

print(my_new_sheet)
#display the id attribute of the my_new_sheet object
my_new_sheet.id


