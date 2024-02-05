import constants


def get_total_classes(sheet):
    """
    Get total number of classes from sheet.
    
    Args:
        - sheet (googleapiclient.discovery.Resource): Authenticated Google Sheets API service object.

    Returns:
        - total_classes (int) : Total number of classes.
    """

    # GET request
    result = (
        sheet.values()
        .get(spreadsheetId=constants.SPREADSHEET_ID, range=constants.TOTAL_CLASSES_CELL)
        .execute()
    )

    # Extract value from result
    value = result.get("values", [])

    # Extract number from string
    total_classes = round(int(value[0][0][26 + 1:].strip()))

    return total_classes


def read_sheet(service):
    """
    Get sheet values.
    
    Args:
        - service (googleapiclient.discovery.Resource): Authenticated Google Sheets API service object.

    Returns:
        - values (list) : 2D list. First dimension represents the sheet's rows.
        - total_classes (int) : Total number of classes.
    """

    # Call the Sheets API
    sheet = service.spreadsheets()

    # GET request for reading the sheet
    result = (
        sheet.values()
        .get(spreadsheetId=constants.SPREADSHEET_ID, range=constants.RANGE_NAME)
        .execute()
    )

    # Extract values from result
    values = result.get("values", [])

    # Get total number of classes
    total_classes = get_total_classes(sheet)

    return values, total_classes


def write_body(values, total_classes):
    """
    Write body of the update request.
    
    Args:
        - values (list) : 2D list containing the sheet's values.
            First dimension represents the sheet's rows.
        
        - total_classes (int) : Total number of classes.

    Returns:
        - body (dict) : body for the update request.
    """

    body = {
       'values': []
    }

    # Minimum attendance to be approved
    absence_limit = 0.25*total_classes

    for i in range(len(values)):
        row = values[i]

        # Average grade for analyzed student
        average = (int(row[2]) + int(row[3]) + int(row[4]))/3

        # Evaluate attendance
        if int(row[1]) <= absence_limit:
            # Evaluate grade
            if average < 50:
                body['values'].append(['Reprovado por Nota', 0])
            elif average < 70:
                naf = round(100 - average)
                body['values'].append(['Exame Final', naf])
            else:
                body['values'].append(['Aprovado', 0])
        else:
            body['values'].append(['Reprovado por Falta', 0])

    return body


def write_state(service):
    """
    Calculate and write results in the sheet.
    
    Args:
        - service (googleapiclient.discovery.Resource): Authenticated Google Sheets API service object.

    Returns:
        None 
    """

    values, total_classes = read_sheet(service)
  
    # Guard clause in case no data is found
    if not values:
        print("No data found.")
        return
  
    # Writing the body of the request, row by row
    body = write_body(values, total_classes)

    # UPDATE request for writing in the sheet
    service.spreadsheets().values().update(
        spreadsheetId=constants.SPREADSHEET_ID,
        range=f'{constants.SHEET_NAME}!G4:H27',
        body=body,
        valueInputOption='RAW'
    ).execute()
