import csv
from datetime import datetime
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference

# Global variable to track the selected spreadsheet/conversion type
selected_spreadsheet = 1

# Inserts a comma-separated string of data into the specified CSV file path using append mode with error handling.
def insertData(file_path, data_string):
    try:
        with open(file_path, "a", newline="") as f:
            f.write(data_string + "\n")
    except Exception as e:
        print(f"Error writing to file: {e}")

# Displays the file path being read and prints the contents of the CSV file using minimal read permissions and error handling.
def viewData(file_path):
    print(f"\n--- Reading file from path: {file_path} ---")
    try:
        with open(file_path, "r") as f:
            contents = f.read()
            if contents.strip() == "":
                print("The file is currently empty.")
            else:
                print(contents)
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' does not exist yet.")
    except Exception as e:
        print(f"Error reading file: {e}")

# Converts numerical data based on the active spreadsheet formula selection.
def convertData(data):
    global selected_spreadsheet
    if selected_spreadsheet == 1:
        # F to C formula: (F - 32) * 5/9
        return (data - 32) * 5 / 9
    elif selected_spreadsheet == 2:
        # lbs to Kg formula: lbs / 2.205
        return data / 2.205
    elif selected_spreadsheet == 3:
        # in to cm formula: in * 2.54
        return data * 2.54
    else:
        return data

# Collects user input entries, calculates conversions, formats data, and stores records into ZooData.csv using insertData.
def getInput():
    try:
        num_entries = int(input("How many entries are being entered? "))
    except ValueError:
        print("Invalid input. Please enter a valid number.")
        return

    for i in range(num_entries):
        print(f"\n--- Entry {i + 1} ---")
        entry_date = input("Enter the date: ")
        
        try:
            entry_value = float(input("Enter the numerical value: "))
        except ValueError:
            print("Invalid numeric value. Skipping entry.")
            continue

        converted_value = convertData(entry_value)
        
        # Format the data into a comma-separated string
        data_record = f"{entry_date},{entry_value},{converted_value:.2f}"

        # Try-except statement to store input into ZooData.csv using insertData
        try:
            insertData("ZooData.csv", data_record)
            print(f"The following data was saved at {datetime.now()}: {data_record}.")
        except Exception as e:
            print(f"Error saving entry to CSV: {e}")

# Function: createChart
# Description: Prompts the user to choose between initial and converted data, reads the CSV, 
#              saves data to 'final.xlsx', and generates a line or bar chart using openpyxl.
# Arguments:
#   - csv_path (str): The file path to the CSV data file.
#   - chart_type (str): The type of chart to generate ('line' or 'bar').
# Returns:
#   - None
def createChart(csv_path, chart_type):
    print("\n--- Select Data Source ---")
    print("1. Initial data")
    print("2. Converted data")
    choice = input("Enter your choice (1 or 2): ").strip()

    while choice not in ["1", "2"]:
        print("Invalid choice. Please enter 1 or 2.")
        choice = input("Enter your choice (1 or 2): ").strip()

    dates = []
    values = []

    try:
        with open(csv_path, mode="r", encoding="utf-8") as file:
            reader = csv.reader(file)
            
            value_col_index = 1 if choice == "1" else 2
            source_name = "Initial Data" if choice == "1" else "Converted Data"

            for row in reader:
                if len(row) > value_col_index:
                    dates.append(row[0])
                    try:
                        # Cast data type to float
                        values.append(float(row[value_col_index]))
                    except ValueError:
                        pass
    except FileNotFoundError:
        print(f"Error: The file '{csv_path}' does not exist yet. Please input data first.")
        return
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        return

    if not values:
        print("No valid data found to generate the chart.")
        return

    # Save data to an Excel spreadsheet titled final.xlsx using openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report Data"

    ws.append(["Date", source_name])
    for d, v in zip(dates, values):
        ws.append([d, v])

    wb.save("final.xlsx")

    # Generate the chart based on the argument passed
    if chart_type.lower() == "line":
        chart = LineChart()
    else:
        chart = BarChart()

    # Set title as <student ID> <current date> (using your student ID banner 'jorkee0238')
    student_id = "jorkee0238"
    current_date = datetime.now().strftime("%m/%d/%Y")
    chart.title = f"{student_id} {current_date}"

    # Label the x and y-axis appropriately
    chart.x_axis.title = "Date"
    chart.y_axis.title = source_name

    # References for data values and categories (dates)
    data_ref = Reference(ws, min_col=2, min_row=1, max_row=len(values) + 1)
    cats_ref = Reference(ws, min_col=1, min_row=2, max_row=len(values) + 1)

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    # Add chart to the worksheet and save workbook
    ws.add_chart(chart, "D2")
    wb.save("final.xlsx")
    print(f"Chart successfully created and saved to 'final.xlsx'!")

# Function: generateReport
# Description: Prompts the user for a graph type ('line' or 'bar') and calls createChart.
# Arguments:
#   - csv_path (str): The file path to the CSV data file.
# Returns:
#   - None
def generateReport(csv_path):
    chart_type = input("Would you like to create a 'line' or 'bar' chart? ").strip().lower()
    while chart_type not in ["line", "bar"]:
        print("Invalid choice. Please enter 'line' or 'bar'.")
        chart_type = input("Would you like to create a 'line' or 'bar' chart? ").strip().lower()

    createChart(csv_path, chart_type)

# Manages the main menu flow control, user option selections, and program routing.
def main():
    print("jorkee0238")
    
    while True:
        print("\n--- Main Menu ---")
        print("1. Input Data")
        print("2. View Current Data")
        print("3. Generate Report")
        print("4. Exit")
        
        choice = input("Enter your menu selection: ").strip()
        now = datetime.now()
        print(f"You chose {choice} at {now}")

        if choice == "1":
            getInput()
        elif choice == "2":
            viewData("ZooData.csv")
        elif choice == "3":
            generateReport("ZooData.csv")
        elif choice == "4":
            print("Exiting program. Goodbye!")
            break
        else:
            print("Invalid choice. Please enter a valid menu option.")


if __name__ == "__main__":
    main()
