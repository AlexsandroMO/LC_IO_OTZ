import pandas as pd

li_base = pd.read_excel('VBA_LIST.xlsm','LI')
li_base = li_base[['TAG', 'TAG_JOIN', 'PREF', 'NAME_INST', 'AREA', 'I/O', 'TAG_CABO_CONTR', 'TAG_CABO_ALIM', 'TIPO_SINAL', 'JBF_CPR', 'JBA_ALIM', 'RULER_JOIN','CH','SLOT','UC','REFERENCIA','FCS','CARD','RULER','HOST','COMP']]
io_create = []
for a in range(len(li_base['TAG'])):
  if li_base['CARD'].loc[a] == 'ALF111-S':
    io_create.append([0,li_base['TAG'].loc[a],li_base['I/O'].loc[a],li_base['CARD'].loc[a],li_base['FCS'].loc[a],'-',li_base['SLOT'].loc[a],li_base['CH'].loc[a],li_base['RULER'].loc[a],li_base['COMP'].loc[a],li_base['HOST'].loc[a],'FOUNDATION FIELDBUS',''])
  else:
    io_create.append([0,li_base['TAG'].loc[a],li_base['I/O'].loc[a],li_base['CARD'].loc[a],li_base['FCS'].loc[a],'-',li_base['SLOT'].loc[a],li_base['CH'].loc[a],li_base['RULER'].loc[a],li_base['COMP'].loc[a],li_base['HOST'].loc[a],'-',''])
  
io_header = ['REV',	'TAG', 'IO',	'CARTAO',	'FCS',	'NODE',	'SLOT','CH',	'REGUA', 'COMP', 	'POS',	'TIPO_SINAL',	'OBS']

df = pd.DataFrame(data=io_create, columns=io_header)
df.to_excel('IO_ALREADY.xlsx')
