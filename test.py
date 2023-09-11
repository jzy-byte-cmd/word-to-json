import docx,json

doc=docx.Document("GKYY097_2022届上海高考英语词汇手册默写P221-226.docx")

data=[]
dic=dict()          #创建字典


for table in doc.tables:            #循环直到遍历所有表格
    for row in table.rows:            #按行遍历表格
            #print(row.cells[0].text,row.cells[1].text,row.cells[3].text)
            dic['id']=row.cells[0].text                #第1行序号，第2行单词，第4行词性和意思
            dic['word']=row.cells[1].text
            dic['meaning']=row.cells[3].text
            #print(dic)
            data.append(dic.copy())                #入栈
            #print(data)
            #input()
            dic.clear()                     #清空字典


doc.save("target.docx")             #保存文件
#print(data)
json_str=json.dumps(data,ensure_ascii=False)    #将列表转为json格式
print(type(json_str))
print(json_str)

file=open("P221-226.json","w+",encoding="utf8")
file.write(json_str)                    #写入文件

file.close()