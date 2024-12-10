import pandas as pd
import numpy as np
import json

def load_excel(path):
  data = pd.read_excel(path,skiprows=[0,1,3,4],na_values=['-'])
  return data

def translate(json_name,dict_json_x):
  if json_name in dict_json_x:
    return dict_json_x[json_name]
  raise ValueError(f'Invalid json name {json_name}!')

def check_num(obj):
  return(isinstance(obj,int) or isinstance(obj,float) or (isinstance(obj,str) and obj.replace('.', '').replace('-','').replace('k','').isnumeric())\
         or obj=='OUT', obj=='ALL IN')

def check_str(obj):
  return(isinstance(obj,str))

def json_maker(obj,df,i,k,v,l_chunk=''):
  if pd.isnull(obj):
    return f'\"{k}\" : null,\n'
  elif v=='num' and check_num(obj):
    return f'\"{k}\" : \"{obj}\",\n'
  elif k=='firstChunk':
    if len(obj)>5:
      return f'\"{k}\" : {int(obj[1:6])},\n'
    else:
      return f'\"{k}\" : {int(obj[1:])},\n'
  elif k=='lastChunk':
    return f'\"{k}\" : {l_chunk},\n'
  elif k=='timeStamp':
    return f'\"{k}\" : \"{obj[5:]}\",\n'
  elif v=='str' and check_str(obj):
    return f'\"{k}\" : \"{obj}\",\n'
  raise ValueError(f'Invalid type of \'{obj}\'! In column \'{df}\'\
               and line {i+1} value type should be \'{v}\' not {type(obj)}!')

def last_ch(k,obj):
  if len(obj)>5:
    return int(obj[-4:])
  else:
    return int(obj[1:])

def template_loop(obj,template,dict_json_x):
  out=""
  l_chunk=np.nan
  for i in range(len(obj)):
    out+='{\n'
    for o in template:
      for sub in template[o]:
        if isinstance(sub,dict):
          out+=f'\"{o}\" : '+'{\n'
          for k,v in sub.items():
            df = translate(k,dict_json_x)
            if k=='firstChunk' and isinstance(obj[df][i],str):
              l_chunk=last_ch(k,obj[df][i])
            out+=json_maker(obj[df][i],df,i,k,v,l_chunk)
          out=out[:-2]+'\n},\n'
        else:
          df = translate(o,dict_json_x)
          out+=json_maker(obj[df][i],df,i,o,sub)
    out=out[:-2]+'},\n'
  return(out.replace('nan','null'))

def custom_to_json(obj):
  template=pd.read_json('template.json', dtype_backend="numpy_nullable")
  with open('dict_json_x.json') as f:
    dict_json_x=json.load(f)
  return ('['+template_loop(obj,template,dict_json_x)[:-2]+'\n]')

def df_to_json(xpath,json_path):
  df = load_excel(xpath)
  json_str=custom_to_json(df)
  print(json_str)
  json_dict = json.loads(json_str)
  print(f'Succefully processed {xpath}!')
  json.dump(json_dict, open(json_path, 'w'), indent=4)
  print(f'JSON file has been generated at {json_path}!')

def test():
  df_to_json('/content/spreadsheets/higs_august_exp_notes.xlsx','test1.json')
test()

