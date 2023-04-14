li_base = pd.read_excel('VBA_LIST.xlsm','LI')
li_base[['TAG', 'TAG_JOIN', 'PREF', 'NAME_INST', 'AREA', 'I/O', 'TAG_CABO_CONTR', 'TAG_CABO_ALIM', 'TIPO_SINAL', 'JBF_CPR', 'JBA_ALIM', 'RULER_JOIN', 'UC', 'REFERENCIA','FCS','CARD','RULER','HOST','COMP']]

df_model = pd.read_excel('VBA_LIST.xlsm','Modelo_INST')


new_list_model = []
for b in range(len(df_model)):
  condutor, service, born_orige, born_dest= [],[],[],[]
  new_sublist = []
  #condutor
  if df_model['CONDUTOR_A'].loc[b] != '-':
    condutor.append(df_model['CONDUTOR_A'].loc[b])
  if df_model['CONDUTOR_B'].loc[b] != '-':
    condutor.append(df_model['CONDUTOR_B'].loc[b])
  if df_model['CONDUTOR_C'].loc[b] != '-':
    condutor.append(df_model['CONDUTOR_C'].loc[b])
  if df_model['CONDUTOR_D'].loc[b] != '-':
    condutor.append(df_model['CONDUTOR_D'].loc[b])
  if df_model['CONDUTOR_E'].loc[b] != '-':
    condutor.append(df_model['CONDUTOR_E'].loc[b])
  #service
  if df_model['SERVICE_A'].loc[b] != '-':
    service.append(df_model['SERVICE_A'].loc[b])
  if df_model['SERVICE_B'].loc[b] != '-':
    service.append(df_model['SERVICE_B'].loc[b])
  if df_model['SERVICE_C'].loc[b] != '-':
    service.append(df_model['SERVICE_C'].loc[b])
  if df_model['SERVICE_D'].loc[b] != '-':
    service.append(df_model['SERVICE_D'].loc[b])
  if df_model['SERVICE_E'].loc[b] != '-':
    service.append(df_model['SERVICE_E'].loc[b])
  #born_orige
  if df_model['BORNE_A_ORIG'].loc[b] != '-':
    born_orige.append(df_model['BORNE_A_ORIG'].loc[b])
  if df_model['BORNE_B_ORIG'].loc[b] != '-':
    born_orige.append(df_model['BORNE_B_ORIG'].loc[b])
  if df_model['BORNE_C_ORIG'].loc[b] != '-':
    born_orige.append(df_model['BORNE_C_ORIG'].loc[b])
  if df_model['BORNE_D_ORIG'].loc[b] != '-':
    born_orige.append(df_model['BORNE_D_ORIG'].loc[b])
  if df_model['BORNE_E_ORIG'].loc[b] != '-':
    born_orige.append(df_model['BORNE_E_ORIG'].loc[b])
  #born_dest
  if df_model['BORNE_A_DEST'].loc[b] != '-':
    born_dest.append(df_model['BORNE_A_DEST'].loc[b])
  if df_model['BORNE_B_DEST'].loc[b] != '-':
    born_dest.append(df_model['BORNE_B_DEST'].loc[b])
  if df_model['BORNE_C_DEST'].loc[b] != '-':
    born_dest.append(df_model['CONDUTOR_C'].loc[b])
  if df_model['BORNE_D_DEST'].loc[b] != '-':
    born_dest.append(df_model['BORNE_D_DEST'].loc[b])
  if df_model['BORNE_E_DEST'].loc[b] != '-':
    born_dest.append(df_model['BORNE_E_DEST'].loc[b])

  #Ajust
  if df_model['TIPO_CABO'].loc[b] != '-':
    for c in range(len(condutor)):
      new_sublist.append([df_model['INST'].loc[b], df_model['TIPO_CABO'].loc[b],condutor[c],service[c],born_orige[c],born_dest[c],df_model['TIPO_INST'].loc[b],'-','-'])
  else:
    new_sublist.append([df_model['INST'].loc[b], df_model['TIPO_CABO'].loc[b],'-','-','-','-',df_model['TIPO_INST'].loc[b],df_model['IO'].loc[b],df_model['LINES'].loc[b]])

  new_list_model.append([df_model['INST'].loc[b], new_sublist])

lc = []
for a in range(len(li_base)):
  for cont in range(len(new_list_model)):
    if li_base['PREF'].loc[a] == new_list_model[cont][0]:
      for b in range(len(new_list_model[cont][1])):
        lc.append([0,li_base['TAG'].loc[a],li_base['TAG_CABO_CONTR'].loc[a],new_list_model[cont][1][b][1],new_list_model[cont][1][b][2],new_list_model[cont][1][b][3],li_base['TAG'].loc[a],new_list_model[cont][1][b][4],('{} | {} | {} / {}').format(li_base['JBF_CPR'].loc[a],li_base['RULER_JOIN'].loc[a],li_base['CH'].loc[a],li_base['SLOT'].loc[a]),new_list_model[cont][1][b][5]])

lc_header = ['REV','TAG_EQUIPAMENTO','TAG_CABO','TIPO','CONDUTOR','TIPO SINAL','ORIGEM','BORNE_ORIGEM','DESTINO','BORNE_DESTINO']

df = pd.DataFrame(data=lc,columns=lc_header)
df.to_excel('LC_ALREADY.xlsx')

new_df = df.groupby(['ORIGEM','TAG_EQUIPAMENTO', 'TAG_CABO', 'DESTINO','TIPO', 'CONDUTOR', 'TIPO SINAL', 'BORNE_ORIGEM', 'BORNE_DESTINO','REV']).count()
new_df.to_excel('LC_CREATE.xlsx')

