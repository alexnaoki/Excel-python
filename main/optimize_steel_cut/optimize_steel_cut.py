import xlwings as xw
from ortools.linear_solver import pywraplp
import pandas as pd

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    codigo_range_i = xw.Range('A3')
    codigo_range_e = codigo_range_i.end('down')

    # print(codigo_range_e.row)
    
    df = codigo_range_i.expand().options(pd.DataFrame).value.reset_index()
    # print(df)
    # print(df['Bitola (mm)'].unique())
    bins_used = 2

    results_row = codigo_range_e.row+bins_used
    for bitola in df['Bitola (mm)'].unique():
        print(f'BITOLA: {bitola} mm')
        xw.Range((results_row-1, codigo_range_e.column)).value = f'BITOLA: {bitola} mm'
        n_list = df.loc[df['Bitola (mm)']==bitola]
        
        weights = []

        for i, n in n_list.iterrows():
            print(f'ROWS: {results_row}')
            

            weights += [float(n['Comprimento (mm)'])]*int(n['Quantidade'])
            
            data = {}
            data['weights'] = weights
            data['items'] = list(range(len(weights)))
            data['bins'] = data['items']
            data['bin_capacity'] = 1200

            run = optimize_steel_cut(data=data)
            print(run)


            print(run[3])
        for i, r in enumerate(run[3]):
            
            xw.Range((results_row, codigo_range_e.column)).value = run[3][i]
            # xw.Range((results_row+i, codigo_range_e.column)).value = run[2][i]
            # bins_used = run[0]

            # results_row += bins_used
            results_row += 1

        results_row += 1


    
def optimize_steel_cut(data):
    print('OPTIMIZING')
    print(data)
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
            sum(x[(i, j)] * data['weights'][i] for i in data['items']) <= y[j] *
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
                        bin_weight += data['weights'][i]
                        bin_pieces.append(data['weights'][i])
                if bin_weight > 0:
                    num_bins += 1
                    all_bin_items.append(bin_items)
                    all_bin_weight.append(bin_weight)
                    all_bin_pieces.append(bin_pieces)
                    print('Bin number', j)
                    print('  Items packed:', bin_items)
                    print('  Total weight:', bin_weight)
                    print()
        print()
        print('Number of bins used:', num_bins)
        print('Time = ', solver.WallTime(), ' milliseconds')
        return (num_bins, all_bin_items, all_bin_weight, all_bin_pieces)
    else:
        print('The problem does not have an optimal solution.')



if __name__ == "__main__":
    xw.Book("optimize_steel_cut.xlsm").set_mock_caller()
    main()
