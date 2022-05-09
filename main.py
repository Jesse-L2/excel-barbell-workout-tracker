"""
main.py
~~~~~~~
The main function of this program uses openpyxl to read the intended weight to be
lifted from various locations in the 'Workout.xlsx' spreadsheet and then apply and
output the weight_calc() function from weight_calculator.py to the various rows and
columns of the 'Workout.xlsx' spreadsheet

Note - 'Workout.xlsx' is a requirement for this program to run, however weight_calc()
will function if called separately onto any int or floating point number, provided that
number is less than the total sum of weights within the weight_calc function in
weight_calculator.py

"""

from openpyxl import Workbook, load_workbook
import weight_calculator


def main():
    try:
        wb = load_workbook('Workout.xlsx', data_only=True)
    # If the workbook doesn't currently exist, create it
    # TODO: create code to build workbook from scratch if it does not exist
    except FileNotFoundError:
        wb = Workbook()
        wb.save(filename="Workout.xlsx")
    # Create variables to reference to particular sheets in the Excel doc
    maxes = wb['Maxes']  # currently unused
    upper_1 = wb['Upper1']
    lower_1 = wb['Lower1']
    upper_2 = wb['Upper2']
    lower_2 = wb['Lower2']
    theo_maxes = wb['Theoretical Weight Scheme']  # currently unused

    """Sample code for updating one single cell
    weight = upper_1['D4'].value # float
    final_weights = weight_calculator.weight_calc(weight)
    final_weights = str(final_weights)
    upper_1['F4'] = final_weights"""

    # Upper1 - Bench Press
    # Update the excel cells, section by section
    # TODO: Update with function to update each row, rather than iterating several times
    for i in range(3, 10):
        bench_total = upper_1.cell(row=i, column=4).value
        upper_1.cell(row=i, column=5).value = str(weight_calculator.weight_calc(bench_total)).replace('[', '').replace(
            ']', '')
    # Upper1 - Overhead Press
    for i in range(3, 9):
        ohp_total = upper_1.cell(row=i, column=10).value
        upper_1.cell(row=i, column=11).value = str(weight_calculator.weight_calc(ohp_total)).replace('[', '').replace(
            ']', '')
    #  Upper2 - Bench Press
    for i in range(3, 8):
        bench_total = upper_2.cell(row=i, column=4).value
        upper_2.cell(row=i, column=5).value = str(weight_calculator.weight_calc(bench_total)).replace('[', '').replace(
            ']', '')
    #  Upper2 - Close-grip Bench Press
    for i in range(11, 15):
        bench_total = upper_2.cell(row=i, column=4).value
        upper_2.cell(row=i, column=5).value = str(weight_calculator.weight_calc(bench_total)).replace('[', '').replace(
            ']', '')

    # Lower Days
    #   Lower1 - Conventional Squat
    for i in range(3, 10):
        squat_total = lower_1.cell(row=i, column=4).value
        lower_1.cell(row=i, column=5).value = str(weight_calculator.weight_calc(squat_total)).replace('[', '').replace(
            ']', '')
    #  Lower1 - Sumo Deadlift
    for i in range(13, 17):
        sdl_total = lower_1.cell(row=i, column=4).value
        lower_1.cell(row=i, column=5).value = str(weight_calculator.weight_calc(sdl_total)).replace('[', '').replace(
            ']', '')
    #  Lower2 - Conventional Deadlift
    for i in range(3, 8):
        cdl_total = lower_2.cell(row=i, column=4).value
        lower_2.cell(row=i, column=5).value = str(weight_calculator.weight_calc(cdl_total)).replace('[', '').replace(
            ']', '')
    #  Lower2 - Front Squat
    for i in range(12, 18):
        squat_total = lower_2.cell(row=i, column=4).value
        lower_2.cell(row=i, column=5).value = str(weight_calculator.weight_calc(squat_total)).replace('[', '').replace(
            ']', '')

    # Save the spreadsheet over the previous spreadsheet
    wb.save('Workout.xlsx')


if __name__ == "__main__":
    main()
