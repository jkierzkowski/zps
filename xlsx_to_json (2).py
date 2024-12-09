import pandas as pd
import numpy as np
import json

def search_between(data,limiting_val,limits):
  return data[data[limiting_val].between(limits[0],limits[1])]
 
def search_exact(data,limiting_val,limits):
  return data[data[limiting_val].isin(limits)]

def load_excel(path):
  data = pd.read_excel(path,skiprows=[0,1,3,4],na_values=['-'])
  return data

def check_num(obj):
  return(isinstance(obj,int) or isinstance(obj,float) or (isinstance(obj,str) and obj.replace('.', '').replace('-','').replace('k','').isnumeric())\
         or obj=='OUT', obj=='ALL IN')

def check_str(obj):
  return isinstance(obj,str)

def json_maker(obj,df,i,k,v,l_chunk=''):
  if pd.isnull(obj):
    return f'\"{k}\" : null,\n'
  elif v=='num' and check_num(obj):
    return f'\"{k}\" : \"{obj}\",\n'
  elif k=='firstChunk':
    if len(obj)>5:
      return f'\"{k}\" : \"{obj[:6]}\",\n'
    else:
      return f'\"{k}\" : \"{obj}\",\n'
  elif k=='lastChunk':
    return f'\"{k}\" : \"{l_chunk}\",\n'
  elif v=='str' and check_str(obj):
    return f'\"{k}\" : \"{obj}\",\n'
  raise ValueError(f'Invalid type of \'{obj}\'! In column \'{df}\'\
               and line {i+1} value type should be \'{v}\' not {type(obj)}!')

def last_ch(k,obj):
  if k=='firstChunk':
    if isinstance(obj,str) and len(obj)>5:
      return obj[-6:]
  return obj

def template_loop(obj,template):
  for i in range(len(obj)):
    out='{\n'
    df_iter=iter(obj)
    df=next(df_iter)
    for o in template:
      for sub in template[o]:
        if isinstance(sub,dict):
          out+=f'\"{o}\" : '+'{\n'
          for k,v in sub.items():
            l_chunk=last_ch(k,obj[df][i])
            out+=json_maker(obj[df][i],df,i,k,v,l_chunk)
            if k!="firstChunk":
              try:
                df = next(df_iter)
              except StopIteration:
                break
          out=out[:-2]+'\n},\n'
        else:
          out+=json_maker(obj[df][i],df,i,o,sub)
          try:
            df = next(df_iter)
          except StopIteration:
            break
    out=out[:-2]+'},\n'
  return(out.replace('nan','null'))

def custom_to_json(obj):
  template  = pd.read_json('template.json', dtype_backend="numpy_nullable")
  return ('['+template_loop(obj,template)[:-2]+'\n]')

def x_to_json(xpath,json_path):
  df = load_excel(xpath)
  json_str=custom_to_json(df)
  json_dict = json.loads(json_str)
  print(f'Succefully processed {xpath}!')
  json.dump(json_dict, open(json_path, 'w'), indent=4)
  print(f'JSON file has been generated at {json_path}!')
def test():
  x_to_json('/content/spreadsheets/higs_august_exp_notes.xlsx','test1.json')
test()

