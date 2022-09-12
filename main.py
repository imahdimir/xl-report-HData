"""

    """

import json
from pathlib import Path

import pandas as pd
from mirutil.df_utils import read_data_according_to_type as read_data
from mirutil.pathes import get_all_subdirs
from mirutil.pathes import has_subdir
from mirutil.df_utils import save_df_as_a_nice_xl as sxl


class Params :
    _pth = '/Users/mahdi/Library/CloudStorage/OneDrive-khatam.ac.ir/Heidari Data/V2'
    root_dir = Path(_pth)

p = Params()

class ColName :
    path = 'path'
    fdir = 'final_dir'
    ls = 'list_dir'
    dp = 'dp'
    mp = 'mp'

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
    foln = 'Folder'
    dsc = 'Description'
    freq = 'Frequency'
    start = 'Start Date'
    end = 'End Date'
    cols = 'Columns'

fc = FinalCol()

def list_dir(dirp: Path) :
    lo = list(dirp.glob('*'))
    lo = [i for i in lo if i.name != '.DS_Store']
    return lo

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
    df[c.fdir] = ~ df[c.path].apply(has_subdir)
    ##
    df = df[df[c.fdir]]
    ##
    df[c.ls] = df[c.path].apply(list_dir)
    ##
    df[c.mp] = df[c.ls].apply(
            lambda x : x[0] if x[0].name == 'meta.json' else x[1]
            )
    ##
    df[c.dp] = df[c.ls].apply(
            lambda x : x[0] if x[0].name != 'meta.json' else x[1]
            )

    ##
    df1 = df.apply(
            lambda x : read_meta(x[c.mp]) , axis = 1 , result_type = 'expand'
            )
    ##
    df = df.join(df1)
    ##
    df = df[[c.path , mc.dcs , mc.freq , mc.start , mc.end , mc.cols]]
    ##
    df[fc.foln] = df[c.path].apply(lambda x : x.relative_to(p.root_dir))
    ##
    df = df[[fc.foln , mc.dcs , mc.freq , mc.start , mc.end , mc.cols]]
    ##
    ren = {
            mc.dcs   : fc.dsc ,
            mc.freq  : fc.freq ,
            mc.start : fc.start ,
            mc.end   : fc.end ,
            mc.cols  : fc.cols ,
            }
    df = df.rename(columns = ren)

    ##
    sxl(df , 'rep.xlsx')

    ##

##
if __name__ == "__main__" :
    main()
    print('Done!')
