#%

#import tkinter
#from tkinter import *
import tkinter as tk

spis_to_find = []

def mycom():
    global spis_to_find
    spis_to_find = edit.get()
    spis_to_find=[x.strip() for x in spis_to_find.split(',')]
#    print(spis_to_find)
#    dir_name = r'C:\Users\17788339\Documents\docs\xlsx_docx'
#    #dir_name = os.getcwd() # если надо упаковать в .exe (только он нормально работает если надо получить путь до самого себя .exe)
#    with open(dir_name + '/_Errors.txt', 'a', encoding = 'cp1251') as Err_txt:
#            Err_txt.write(spis_to_find + '\n') 

def main():
    from docx import Document
    from os import listdir
    from pandas import ExcelFile
    from pandas import DataFrame
#    from pandas import to_numeric

    global spis_to_find
    print(spis_to_find)
    
    # dir_name = r'C:\Users\_docs\xlsx_docx'
    # dir_name = os.getcwd() # если надо упаковать в .exe (только он нормально работает если надо получить путь до самого себя .exe)
    # dir_name = path.abspath("mydir/example.txt") 
    
#    spis_to_find = ['азмер расчетного резерва','орматив резерва'] # список ключевых фраз по которым тянуть values
    
    Data_xlsx = []
    key_conteins = []
    Err_files = []
    
    all_docx_xlsx = [f for f in listdir(dir_name) if f.endswith('.docx')] + [f for f in listdir(dir_name) if f.endswith('.xlsx')]
    #print(all_docx_xlsx)
    #print(all_docx_xlsx)
    print('\n' + str(len(all_docx_xlsx))+' - All .xlsx & .docx files in dir_name ' + '\n')
    list_global = []
    list_global.append(all_docx_xlsx)
    #print(list_global)
    #print(len(list_global))
    try:
        
        for xlsx_name in [f for f in listdir(dir_name) if f.endswith('.xlsx')]:
            
            try: 
                
                xl = ExcelFile(dir_name + '/' + xlsx_name)
                df1 = xl.parse()
                #print (df1)
                
            except Exception as Exc_xlsx:
    #            print(str(Exc_xlsx) + ' \t ' + str(xlsx_name) + '\n') 
    #            with open(dir_name + '/_Errors.txt', 'a', encoding = 'cp1251') as Err_txt:
    #                Err_txt.write(str(Exc_xlsx) + ' \t ' + str(xlsx_name) + '\n') 
                Err_files.append(str(xlsx_name) + '    Exception: ' + str(Exc_xlsx) + '\n') 
        
            my_columns = df1.columns
            for col in my_columns:
                my_val_ind = df1[df1[col].astype(str).apply(lambda x: any(s in x for s in spis_to_find))].index
                # ищем вхождение ключевой фразы из списка spis_to_find (либо функц any) и берем индекс
        #            print(my_val_ind)
                if len(my_val_ind ) > 0: # если индекс какой то был найден то тянем значения по первому вхождению
                    for ival in df1[df1.index == my_val_ind[0]].values:
                        Data_xlsx.append(xlsx_name)
                        key_conteins.append(ival)
    #                    print(ival)
    #                    print(xlsx_name) # список имен файлов в пути dir_name
    #    print(Data_xlsx) # список обработанных .xlsx и поличество
    #    print(len(Data_xlsx))               
    #    spis_xlsx = {}
    #    spis_xlsx1 = []
        spis_xlsx2 = []
        spis_xlsx3 = []
        spis_xlsx4 = []
        for XLSX in range(len(Data_xlsx)):
    #        spis_xlsx.update({(str(XLSX)+'_XLSX'):[Data_xlsx[XLSX], key_conteins[XLSX][0], key_conteins[XLSX][1]]})
    #        spis_xlsx1.append([key_conteins[XLSX][0], key_conteins[XLSX][1], Data_xlsx[XLSX]])
            spis_xlsx2.append(key_conteins[XLSX][0])
            spis_xlsx3.append(key_conteins[XLSX][1])
            spis_xlsx4.append(Data_xlsx[XLSX])
        #        print(XLSX)
    #    print(spis_xlsx)
    #    print(spis_xlsx1)
        
    #    print(spis_xlsx2)
    #    print(len(spis_xlsx2))
    #    print(len(spis_xlsx3))
    #    print(len(spis_xlsx4))
    
        ##########################################################################################
            
        Data_DOC = []
        key_conteins_doc = []
        for docx_name in [f for f in listdir(dir_name) if f.endswith('.docx')]:
            
            try:
                
                document = Document(dir_name + '/' + docx_name)
                table = document.tables[0]
                
            except Exception as Exc_docx:
    #            print(str(Exc_docx) + ' \t ' + str(docx_name) + '\n') # 
    #            with open(dir_name + '/_Errors.txt', 'a', encoding = 'cp1251') as Err_txt:
    #                Err_txt.write(str(Exc_docx) + ' \t ' + str(docx_name) + '\n')
                Err_files.append(str(docx_name) + '    Exception: ' + str(Exc_docx) + '\n') 
                
            data = []
            for i, row in enumerate(table.rows):
                 text = (cell.text for cell in row.cells)
                 if i == 0:
                    header = tuple(text)
                    #print(keys)
                    continue
                 row_data = tuple(text)
                 data.append(row_data)
                 print(row_data)
            df2 = DataFrame(data,columns=header)
            #print (df2)
            docx_columns = df2.columns
            for col in docx_columns:
                    my_val_ind = df2[df2[col].astype(str).apply(lambda x: any(s in x for s in spis_to_find))].index
        #                print(my_val_ind)
                    if len(my_val_ind ) > 0:
                        for ival in df2[df2.index == my_val_ind[0]].values:
                            Data_DOC.append(docx_name)
                            key_conteins_doc.append(ival)
        #                        print(ival)
        #                    print(docx_name)
    #    print(Data_DOC) # список обработанных .docx и поличество
    #    print(len(Data_DOC))
    #    spis_docx = {}
    #    spis_docx1 = []
        spis_docx2 = []
        spis_docx3 = []
        spis_docx4 = []
        for DOC in range(len(Data_DOC)):
    #        spis_docx.update({(str(DOC)+'_DOC'):[Data_DOC[DOC], key_conteins_doc[DOC][0], key_conteins_doc[DOC][1]]})
    #        spis_docx1.append([key_conteins_doc[DOC][0], key_conteins_doc[DOC][1], Data_DOC[DOC]])
    #    print(spis_docx1)
            spis_docx2.append(key_conteins_doc[DOC][0])
            spis_docx3.append(key_conteins_doc[DOC][1])
            spis_docx4.append(Data_DOC[DOC])
    #    0 - тут просто глобальный список потом к нему все аппендить и в df )
    #    print(spis_docx2)
    #    print(len(spis_docx2))
    #    print(len(spis_docx3))
    #    print(len(spis_docx4))
        s2 = spis_xlsx2 + spis_docx2
        s3 = spis_xlsx3 + spis_docx3
        s4 = spis_xlsx4 + spis_docx4
    #    print(s4)
    #    print(len(s4))
        list_global.extend((s2,s3,s4,Err_files))
    #    list_global.append(Err_files)
        NOT_read = [item for item in all_docx_xlsx if item not in frozenset(s4)] 
        print(all_docx_xlsx)
        print(len(all_docx_xlsx))
        print(s4)
        print(len(s4))
        print(NOT_read)
        print(len(NOT_read))
        
        list_global.append(NOT_read)
    #    print(list_global)
    #    print(len(list_global))
        
        df_list_global = DataFrame(list_global).T
        df_list_global.columns = ['Все файлы .docx .xlsx в папке','Вариант ключевых фраз','Значение','Обработанные файлы','Ошибки в обработанных файлах','НЕ обработанные']
#        to_numeric(df_list_global['Значение'])
    
    #    print(df_list_global)
    #    print(len(df_list_global))
    #    df_list_global.sort_values(by=['Все файлы .docx .xlsx в папке','Ошибки в обработанных файлах','НЕ обработанные'], ascending=False, na_position='first')
    
    
    ##    1 - тут если все через DataFrame и  merge клеить dfы )) 
    
    #    df_all_files = DataFrame(all_docx_xlsx)
    #    df_all_files.columns = ['Все файлы .docx .xlsx в папке']
    ##    print(df_all_files)
    #    df_read_files = DataFrame(spis_xlsx1 + spis_docx1)
    #    df_read_files.columns = ['Вариант ключевых фраз','Значение','Обработанные файлы']
    ##    print(df_read_files)
    #    df_m1 = df_all_files.merge(df_read_files, left_on = ['Все файлы .docx .xlsx в папке'], right_on = ['Обработанные файлы'], how = 'outer' )
    ##    df_m1.to_excel(dir_name + '\_DF.xlsx', encoding = 'cp1251', index = False)    
    #    df_Err_files = DataFrame(Err_files)
    #    df_Err_files.columns = ['Ошибки в обработанных файлах']
    ##    print(df_Err_files)
    #    NOT_read = [item for item in [f for f in listdir(dir_name)] if item not in frozenset(Data_xlsx+Data_DOC)]
    
    #    df_m2 = df_m1.merge(df_Err_files, left_on = ['Все файлы .docx .xlsx в папке'], right_on = ['Ошибки в обработанных файлах'], how = 'outer' )
    ##    print(df_m2)
    ##    df_m2.to_excel(dir_name + '\_DF.xlsx', encoding = 'cp1251', index = False) 
    #    
    #    df_NOT_read = DataFrame([item for item in [f for f in listdir(dir_name)] if item not in frozenset(Data_xlsx+Data_DOC)])
    #    df_NOT_read.columns = ['НЕ обработанные']
    ##    print(df_NOT_read)
    ##    print(len(df_NOT_read))
    #    
    #    df_m3 = df_m2.merge(df_NOT_read, left_on = ['Все файлы .docx .xlsx в папке'], right_on = ['НЕ обработанные'], how = 'outer' )
    #    print(df_m3)
    
    
    ##    2 - тут к df добалял списки (с той же размерностью что и df иначе ошибка Exception: Length of values does not match length of index)
    #    aa = []
    #    for f_err in range(len(all_docx_xlsx)-len(Err_files)):
    #        aa.append('')
    ##    print(aa)
    ##    print(len(aa))
    #    Err_files = Err_files + aa
    ##    print(Err_files)
    ##    print(len(Err_files))
    #    df_m1['Ошибки в обработанных файлах'] = Err_files
    ##    print(df_m1)
    #    bb = []
    #    for f_notread in range(len(all_docx_xlsx)-len(NOT_read)):
    #        bb.append('')
    ##    print(bb)
    ##    print(len(bb))
    #    NOT_read = NOT_read + bb
    ##    print(NOT_read)
    ##    print(len(NOT_read))
    #    df_m1['НЕ обработанные'] = NOT_read
    #    print(df_m1)
    ##    df_m1.to_excel(dir_name + '\_DF.xlsx', encoding = 'cp1251') 
    #    df_m1.to_excel(dir_name + '\_DATA.xlsx', encoding = 'cp1251', index = False) 
    #    df_m1.to_csv(dir_name + '\_DATA.csv', encoding = 'cp1251', index = False)    
    
    #    print(str(len(all_docx_xlsx)) + '  All  .xlsx & .docx in dir_name ')
    #    print(str(len(Err_files)) + '  Err_files ')
    #    print(str(len(NOT_read)) + '  NOT_read ')
    #    print(str(len(all_docx_xlsx)-len(Err_files)) + '  all_docx_xlsx - Err_files ')
        
    # извращался с сортингом дропингом и тому подобное что бы убрать после мёрджа пустые яч
    #    df_m3['НЕ обработанные'].replace('', np.nan, inplace=True)
    #    df_m3.dropna(subset=['НЕ обработанные'], inplace=True)
    #    df_m3.sort_values(by=['Все файлы .docx .xlsx в папке', 'Ошибки в обработанных файлах', 'НЕ обработанные'], ascending=False, na_position='last')
    #    df_m3.reset_index(level='Ошибки в обработанных файлах')
    #    df_m3 = df_m3[~df_m3['Ошибки в обработанных файлах'].isnull()]
    #    df_m3.to_excel(dir_name + '\_DF.xlsx', encoding = 'cp1251', index = False) 
        
    except Exception as Exc_all: # ЕСЛИ err то этот блок если нет то нет
        print(str(xlsx_name +'//'+ docx_name  + '    Exception: ' + str(Exc_all)) + '\n')  # напечатать ошибку и Обработанные файлы .xlsx на котором была она вызвана
        with open(dir_name + '/_Errors.txt', 'a', encoding = 'cp1251') as Err_txt:
            Err_txt.write(str(Exc_all) + ' \t ' + str(xlsx_name + '/' + docx_name) + '\n') 
            
    df_list_global.to_excel(dir_name + '\_DATA.xlsx', encoding = 'cp1251', index = False) 
    df_list_global.to_csv(dir_name + '\_DATA.csv', encoding = 'cp1251', index = False)    
    
    # Варианты 
    # сохранение в формат xlsx
    #writer = ExcelWriter(dir_name + '\_data.xlsx', engine='xlsxwriter')
    #df_read_files.to_excel(writer, 'Sheet1')
    #writer.save()
    # сохранение в формат csv 
    #df_read_files.to_csv(dir_name + '\_data.csv', encoding = 'utf-8')    
    #df_read_files.to_csv(dir_name + '\_data.csv', encoding = 'utf-16')    
    #df_read_files.to_csv(dir_name + '\_data.csv', encoding = 'utf_16_sig')    
    #df_read_files.to_csv(dir_name + '\_data.csv', encoding = 'cp1251')

readme_text = '\
Файл Pars_DOCX+XLSX(Tkinter_exe).py должен лежать вместе с теми \n \
файлами (.xlsx .docx) в папке которые парсит (с которыми работает). \n \
если .doc .rtf - воспользоваться конверторами если .pdf - (ABBY FineReader) \n \
\n \
В окно для ввода пишем все варианты ключевых фраз в поле, по которым \n \
нужна информация (ключевые фразы могут быть разными в разных документах). \n \
\n \
Ключевые фраз должны быть записаны следующим образом: \n \
Размер расчетного резерва, Норматив резерва\n \
(можно так – азмер расчетного резерв, орматив резер). \n \
\n \
В результате будут сформированы файлы _Errors.txt с ошибками если они будут, \n \
Excel (_DATA.xlsx, _DATA.csv) в этой же папке со столбцами: \n \
Все файлы .docx .xlsx в папке; Вариант ключевых фраз	; Значение; \n \
Обработанные файлы; Ошибки в обработанных файлах; НЕ обработанные.\
'

win = tk.Tk() # большая Tk иначе ошибка 'module' object is not callable
win.title("Pars_DOCX+XLSX.py (GUI на Python)")
win.geometry('990x800')

#    ИНФА РИДМИ readme (заголовок)
t1 = tk.Label(win, text = 'README (ИНСТРУКЦИЯ)', font="Algerian 20", fg = '#9b2d30')
#t1.config(font = ('Courier', 25)) #
t1.grid(row=0, column=0, padx=5, pady=5, sticky="s")

#    ИНФА РИДМИ readme (сам текст)
t1 = tk.Label(win, text = readme_text, font=("Bad Script", 12, "bold"), fg = '#9b2d30')
t1.grid(row=1, column=0, padx=5, pady=5, sticky="s")

#    ПОЛЕ ВВОДА
edit = tk.Entry (win, width = 45, font=("Bad Script", 12, "bold"), bg = '#d5ffd7', bd='3')
edit.grid(row=2, column=0, padx=10, pady=10, sticky="s")

#    КНОПКА ПРИМЕНИТЬ СПИСОК
button_1 = tk.Button(win, height=4, width=50, text= '#1 ИСПОЛЬЗОВАТЬ ЭТОТ СПИСОК',    \
                      bg='#9b2d30', fg="#ffffff", bd='3', font=("Courier", 11, "bold"), command=mycom)
button_1.grid(row=3, sticky="s")

#    КНОПКА ПАРСИТЬ ДОКУМЕНТЫ И СФОРМИРОВАТЬ НАЙДЕННУЮ ИНФОРМАЦИЮ В ТАБЛИЦИ xlsx и csv
button_2 = tk.Button(win, height=2, width=100, text= '#2 ПАРСИТЬ ДОКУМЕНТЫ И СФОРМИРОВАТЬ НАЙДЕННУЮ ИНФОРМАЦИЮ В ТАБЛИЦИ xlsx и csv',    \
                      bg='#9b2d30', fg="#ffffff", bd='3', font=("Courier", 11, "bold"), command=main)
button_2.grid(row=5, column=0, padx=35, pady=35, sticky="s")


win.mainloop()



