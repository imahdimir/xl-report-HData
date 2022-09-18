"""

    """

import json
from pathlib import Path

import pandas as pd
from mirutil.pathes import get_all_subdirs
from mirutil.pathes import has_subdir


class Params :
    _pth = '/Users/mahdi/Library/CloudStorage/OneDrive-khatam.ac.ir/Datasets/Heidari Data/V2'
    root_dir = Path(_pth)

p = Params()

class ColName :
    path = 'path'
    rpath = 'rpath'
    max_lvl = 'max_lvl'
    mfp = 'mfp'

c = ColName()

class MetaCol :
    dcs = 'dcs'
    freq = 'freq'
    start = 'start'
    end = 'end'
    tcol = 'tcol'
    cols = 'cols'

mc = MetaCol()

class FinalCol :
    foln = 'Folder Name'
    dsc = 'Short Description'
    freq = 'Frequency'
    start = 'Start Date'
    end = 'End Date'
    cols = 'Columns'
    _1st_lvl = '1st Level Category'
    _2nd_lvl = '2nd Level'

fc = FinalCol()

def read_meta(mp) :
    with open(mp) as f :
        return json.load(f)

def main() :
    pass

    ##
    subs = get_all_subdirs(p.root_dir)
    ##
    df = pd.DataFrame()
    df[c.path] = list(subs)
    ##
    msk = ~ df[c.path].apply(has_subdir)

    df = df[msk]
    ##
    df[c.rpath] = df[c.path].apply(lambda x : x.relative_to(p.root_dir))
    ##
    df[c.max_lvl] = df[c.rpath].apply(lambda x : len(x.parts) - 1)
    ##
    max_all = df[c.max_lvl].max()
    max_all_dc = {
            'max' : str(df[c.max_lvl].max())
            }
    with open('max.json' , 'w') as f :
        json.dump(max_all_dc , f , indent = 4)
    ##
    for i in range(max_all) :
        fu = lambda x : x.parts[i] if len(x.parts) - 1 > i else None
        df[i + 1] = df[c.rpath].apply(fu)

    ##
    df[fc.foln] = df[c.rpath].apply(lambda x : x.name)
    ##
    df[c.mfp] = df[c.path] / 'meta.json'
    ##
    _fu = lambda x : read_meta(x[c.mfp])
    df1 = df.apply(_fu , axis = 1 , result_type = 'expand')
    ##
    df = df.join(df1)
    ##
    cols = list(range(1 , max_all))
    cols += [fc.foln , mc.dcs , mc.freq , mc.start , mc.end]
    df = df[cols]
    ##
    df = df.sort_values(list(range(1 , max_all)) + [fc.foln])
    ##
    ren = {
            1        : fc._1st_lvl ,
            2        : fc._2nd_lvl ,
            mc.dcs   : fc.dsc ,
            mc.freq  : fc.freq ,
            mc.start : fc.start ,
            mc.end   : fc.end ,
            }
    df = df.rename(columns = ren)
    ##
    for col in [fc.start , fc.end] :
        df[col] = df[col].str.replace('-$' , '')
    ##
    df.to_excel('update.xlsx' , index = False)

    ##

##
if __name__ == "__main__" :
    main()
    print('Done!')
