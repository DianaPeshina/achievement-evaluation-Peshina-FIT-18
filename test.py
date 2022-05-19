
import math
import pandas as pd
data = pd.read_excel(r'C:\Users\Computer\Desktop\ДИПЛОМ\рейтинг успеваемости хороших людей.xlsx',0)

credit_map = {'незачет': 2, 'зачет': 3}
data=data.replace(credit_map)

goal = data.loc[0]['экзамен1':].values.tolist()
newdf=data[1:].copy()

def rulew(vector, goal):
    goal
    result = 0
    ab=0
    a2=0
    b2=0
    size=len(goal)
    for i in range(size):
        ab=ab+goal[i]*vector[i]
        a2=a2+math.pow(goal[i],2)
        b2=b2+math.pow(vector[i],2)
    result=ab/(math.sqrt(a2)*math.sqrt(b2)) 
    return result

def ruleb(vector, goal):
    goal
    result = 0
    ab=0
    a2=0
    size=len(goal)
    for i in range(size):
        ab=ab+goal[i]*vector[i]
        a2=a2+math.pow(goal[i],2)
    result=ab*100/a2 
    return result

newdf['w'] = data.apply(lambda x: rulew(x.loc['экзамен1':].values.tolist(), goal), axis =  1)
newdf['b'] = data.apply(lambda x: ruleb(x.loc['экзамен1':].values.tolist(), goal), axis =  1)
print(str(newdf.sort_values(['b','w'],ascending=[False, False])))