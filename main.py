from openpyxl import Workbook, load_workbook
import weight_calculator


def main():
    try:
        wb = load_workbook('Workout.xlsx', data_only=True)
    except FileNotFoundError:
        wb = Workbook()
        wb.save(filename="Workout.xlsx")

    maxes = wb['Maxes']
    upper_1 = wb['Upper1']
    lower_1 = wb['Lower1']
    upper_2 = wb['Upper2']
    lower_2 = wb['Lower2']
    theo_maxes = wb['Theoretical Weight Scheme']

    '''Sample code for updating one single cell
    weight = upper_1['D4'].value # float
    final_weights = weight_calculator.weight_calc(weight)
    final_weights = str(final_weights)
    upper_1['F4'] = final_weights'''

    # Bench
    # Upper1-BP
    # Updating the excel cells, section by section
    # Note: this code is not as efficient as it could be
    # TODO: Update with function to update each row
    for i in range(3, 10):
        bench_total = upper_1.cell(row=i, column=4).value
        upper_1.cell(row=i, column=5).value = str(weight_calculator.weight_calc(bench_total)).replace('[', '').replace(
            ']', '')
    # OHP -D1
    for i in range(3, 9):
        ohp_total = upper_1.cell(row=i, column=10).value
        upper_1.cell(row=i, column=11).value = str(weight_calculator.weight_calc(ohp_total)).replace('[', '').replace(
            ']', '')
    #  Upper2-BP
    for i in range(3, 8):
        bench_total = upper_2.cell(row=i, column=4).value
        upper_2.cell(row=i, column=5).value = str(weight_calculator.weight_calc(bench_total)).replace('[', '').replace(
            ']', '')
    #  Upper2-CG Bench
    for i in range(11, 15):
        bench_total = upper_2.cell(row=i, column=4).value
        upper_2.cell(row=i, column=5).value = str(weight_calculator.weight_calc(bench_total)).replace('[', '').replace(
            ']', '')

    # Lower Days
    #   Lower1-S
    for i in range(3, 10):
        squat_total = lower_1.cell(row=i, column=4).value
        lower_1.cell(row=i, column=5).value = str(weight_calculator.weight_calc(squat_total)).replace('[', '').replace(
            ']', '')
    #  Lower1-SDL
    for i in range(13, 17):
        sdl_total = lower_1.cell(row=i, column=4).value
        lower_1.cell(row=i, column=5).value = str(weight_calculator.weight_calc(sdl_total)).replace('[', '').replace(
            ']', '')
    #  Lower2-CDL
    for i in range(3, 8):
        cdl_total = lower_2.cell(row=i, column=4).value
        lower_2.cell(row=i, column=5).value = str(weight_calculator.weight_calc(cdl_total)).replace('[', '').replace(
            ']', '')
    #  Lower2-FS
    for i in range(12, 18):
        squat_total = lower_2.cell(row=i, column=4).value
        lower_2.cell(row=i, column=5).value = str(weight_calculator.weight_calc(squat_total)).replace('[', '').replace(
            ']', '')


    wb.save('Workout.xlsx')


if __name__ == "__main__":
    main()
