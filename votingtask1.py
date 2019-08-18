from xlrd import open_workbook
from xlutils.copy import copy
b=open_workbook('votinginfo.xls')
wb=copy(b)
flag=0
ws=wb.get_sheet(0)
fs=b.sheet_by_index(0)
p=[0,0,0,0,0]
passwordp='votepass'
exitpass='exitvote'
namec=fs.col_values(1)
print(namec)
while(1):
    print('''cast your vote to:
                 for candidate 1:press 1
                 for candidate 2:press 2
                 for candidate 3:press 3
                 for candidate 4:press 4
                 for candidate 5:press 5
                 for result press 6
                 for exit press 7''')
    v=int(input('enter your choice'))
    if v==1:
        p[0]+=1
    if v==2:
        p[1]+=1
    if v==3:
        p[2]+=1
    if v==4:
        p[3]+=1
    if v==5:
        p[4]+=1
    ws.write(1,2,p[0])
    ws.write(2,2,p[1])
    ws.write(3,2,p[2])
    ws.write(4,2,p[3])
    ws.write(5,2,p[4])
    if v==6:
        att=3
        while att>0:
            pa=input('enter password')
            if pa==passwordp:
                print('the winner is candidate',end=' ')
                print(p.index(max(p))+1)
                for i in range(0,5):
                    print('votes of {} is {}'.format(namec[i+1],p[i]))
                break
            else:
                print('you are left with {} attempts'.format(att-1))
                att-=1
        if att==0:
            print('attempts failed')
    if v==7:
        at=3
        while at>0:
            pa=input('enter password')
            if pa==exitpass:
                flag=1
                break
            else:
                print('wrong password....try again in %s attempts'%(at-1))
        if flag==1:
            break
wb.save('votinginfo.xls')
