import pandas as pd
import numpy as np
import json

def load_excel(path):
  data = pd.read_excel(path,skiprows=[0,1,3,4],na_values=['-'])
  return data

def check_num(obj):
  return(isinstance(obj,int) or isinstance(obj,float) or pd.isnull(obj) or (isinstance(obj,str) and obj.replace('.', '').replace('-','').replace('k','').isnumeric())\
         or obj=='OUT', obj=='ALL IN')

def check_str(obj):
  return(isinstance(obj,str) or pd.isnull(obj))

def custom_to_json(obj):
  template  = pd.read_json('template.json')
  out = '[\n'
  for i in range(len(obj)-1):
    out+='{\n'
    df_iter = iter(obj)
    df = next(df_iter)
    for o in template:
      x=''
      for sub in template[o]:
        if isinstance(sub,dict):
          out+=f'\"{o}\" : '+'{\n'
          for k,v in sub.items():
            if v=='num' and check_num(obj[df][i]):
              out+=f'\"{k}\" : \"{obj[df][i]}\",\n'
            elif k=='firstChunk':
              if obj[df][i]==r'_\d{4} to _\d{4}':
                out+=f'\"{k}\" : \"{obj[df][i][:5]}\",\n'
                x = obj[df][i][-5:]
              else:
                out+=f'\"{k}\" : \"{obj[df][i]}\",\n'
                x=obj[df][i]
            elif k=='lastChunk':
                out+=f'\"{k}\" : \"{x}\",\n'
            elif v=='str' and check_str(obj[df][i]):
              out+=f'\"{k}\" : \"{obj[df][i]}\",\n'
            else:
              raise ValueError(f'Invalid type of \'{obj[df][i]}\'! In column \'{df}\' and line {i+1} value type should be \'{v}\' not {type(obj[df][i])}!')
            if k!='firstChunk':
              try:
                df = next(df_iter)
              except StopIteration:
                break
          out=out[:-2]
          out+='\n},\n'
        else:
          if sub=='num' and check_num(obj[df][i]):
            out+=f'\"{o}\" : {obj[df][i]},\n'
          elif sub=='str' and check_str(obj[df][i]):
            out+=f'\"{o}\" : \"{obj[df][i]}\",\n'
          else:
            raise ValueError(f'Invalid type of \'{obj[df][i]}\'! In object \'{o}\' value type should be {sub}!')
          try:
            df = next(df_iter)
          except StopIteration:
            break
    out=out[:-2]
    out+='},\n'
  out=out[:-2]+'\n]'
  out = out.replace('nan','null')
  return out

def df_to_json(xpath,json_path):
  df = load_excel(xpath)
  json_str=custom_to_json(df)
  json_dict = json.loads(json_str)
  print(f'Succefully processed {xpath}!')
  json.dump(json_dict, open(json_path, 'w'), indent=4)
  print(f'JSON file has been generated at {json_path}!')
  
def test():
  df_to_json('higs_august_exp_notes.xlsx','test1.json')
test()