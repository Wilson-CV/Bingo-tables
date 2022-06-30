import pandas as pd
import random as rd

"""
Formato de tablas 
columna “B” números entre el 1 y el 15, 
columna “I” números del 16 al 30, 
columna “N” números del 31 al 45, 
columna “G” números 46 al 60 
columna “O” números del 61 al 75
"""

n_tablas = 2300
grid_size = 5

B = range(1, 16)
I = range(16, 31)
N = range(31, 46)
G = range(46, 61)
O = range(61, 76)

writer = pd.ExcelWriter('tablas.xlsx')

col = 0
row = 0
sheet_num = 1
spaces_r = 1
spaces_c = 4

for i in range(n_tablas):
    card = {'B': (rd.sample(B, grid_size)),
            'I': (rd.sample(I, grid_size)),
            'N': (rd.sample(N, grid_size)),
            'G': (rd.sample(G, grid_size)),
            'O': (rd.sample(O, grid_size))}
    df = pd.DataFrame(card, columns=['B', 'I', 'N', 'G', 'O'])
    # create excel writer object
    # write dataframe to excel
    if(i != 0):
        if(((i % 6) == 0)):
            row = row + grid_size + 3 + 1
        else:
            row = row + grid_size + spaces_r + 1
    df.to_excel(writer, sheet_name='Sheet'+str(sheet_num),
                index=False, header=True, startrow=row, startcol=col)


# save the excel
writer.save()


print('DataFrame is written successfully to Excel File.')
