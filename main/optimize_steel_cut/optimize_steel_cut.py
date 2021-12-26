import xlwings as xw
from ortools.linear_solver import pywraplp
import pandas as pd

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    print('Reading Input')
    inputTable_Range = xw.Range('A3')
    inputTable_Range_end = inputTable_Range.end('down')
    
    print('Creating DataFrame')
    df_inputTable = inputTable_Range.expand().options(pd.DataFrame).value.reset_index()

    results_row = inputTable_Range_end.row+3

    results_Range = xw.Range((results_row, inputTable_Range.column))
    print('Clearing previous result...')
    clear_range_e = xw.Range((results_row, inputTable_Range.column)).end('down')
    clear_range_e = clear_range_e.end('right')
    # clear_range_e.color = (255,0,0)
    xw.Range(results_Range, clear_range_e).clear()
    print('Done Clearing')
    
    print('Optimizing:')
    for bitola in df_inputTable['Bitola (mm)'].unique():
        print(f'BITOLA: {bitola} mm')
        xw.Range((results_row-1, inputTable_Range_end.column)).value = f'BITOLA: {bitola} mm'
        xw.Range((results_row, inputTable_Range_end.column)).value = [['Used', 'Wasted']]
        results_row += 1

        df_bitola = df_inputTable.loc[df_inputTable['Bitola (mm)']==bitola]
        
        # List all segments for a particular diameter
        segments = []
        for i, n in df_bitola.iterrows():
            # print(f'ROWS: {results_row}')          
            segments += [float(n['Comprimento (cm)'])]*int(n['Quantidade'])
            
        data = {}
        data['segments'] = segments
        data['items'] = list(range(len(segments)))
        data['bins'] = data['items']
        data['bin_capacity'] = 1200

        # Optimize function
        run = optimize_steel_cut(data=data)

        # Display results
        for i, r in enumerate(run[3]):
            # List of segments for 1 1200cm rebar
            xw.Range((results_row, inputTable_Range_end.column+2)).value = run[3][i]
            
            # Used and Wasted 1200cm rebar
            xw.Range((results_row, inputTable_Range_end.column)).value = run[2][i]      
            xw.Range((results_row, inputTable_Range_end.column+1)).value = 1200-run[2][i]

            xw.Range((results_row, inputTable_Range_end.column)).color = (0,255,0)      # Green
            xw.Range((results_row, inputTable_Range_end.column+1)).color = (255,0,0)    # Red

            results_row += 1

        results_row += 1
        print('Done')


    
def optimize_steel_cut(data):
    # print('OPTIMIZING')
    # print(data)
    solver = pywraplp.Solver.CreateSolver('SCIP')
    
    # Variables
    # x[i, j] = 1 if item i is packed in bin j.
    x = {}
    for i in data['items']:
        for j in data['bins']:
            x[(i, j)] = solver.IntVar(0, 1, 'x_%i_%i' % (i, j))

    # y[j] = 1 if bin j is used.
    y = {}
    for j in data['bins']:
        y[j] = solver.IntVar(0, 1, 'y[%i]' % j)

    # Constraints
    # Each item must be in exactly one bin.
    for i in data['items']:
        solver.Add(sum(x[i, j] for j in data['bins']) == 1)

    # The amount packed in each bin cannot exceed its capacity.
    for j in data['bins']:
        solver.Add(
            sum(x[(i, j)] * data['segments'][i] for i in data['items']) <= y[j] *
            data['bin_capacity'])

    # Objective: minimize the number of bins used.
    solver.Minimize(solver.Sum([y[j] for j in data['bins']]))

    status = solver.Solve()

    
    if status == pywraplp.Solver.OPTIMAL:
        num_bins = 0.
        all_bin_items = []
        all_bin_weight = []
        all_bin_pieces = []

        for j in data['bins']:
            if y[j].solution_value() == 1:
                bin_items = []
                bin_pieces = []
                bin_weight = 0
                for i in data['items']:
                    if x[i, j].solution_value() > 0:
                        bin_items.append(i)
                        bin_weight += data['segments'][i]
                        bin_pieces.append(data['segments'][i])
                if bin_weight > 0:
                    num_bins += 1
                    all_bin_items.append(bin_items)
                    all_bin_weight.append(bin_weight)
                    all_bin_pieces.append(bin_pieces)
        #             print('Bin number', j)
        #             print('  Items packed:', bin_items)
        #             print('  Total weight:', bin_weight)
        #             print()
        # print()
        # print('Number of bins used:', num_bins)
        # print('Time = ', solver.WallTime(), ' milliseconds')
        return (num_bins, all_bin_items, all_bin_weight, all_bin_pieces)
    else:
        print('The problem does not have an optimal solution.')
        
if __name__ == "__main__":
    xw.Book("optimize_steel_cut.xlsm").set_mock_caller()
    main()
