"""

    """

import json
from pathlib import Path

import pandas as pd
from mirutil.df_utils import read_data_according_to_type as read_data
from mirutil.pathes import get_all_subdirs
from mirutil.pathes import has_subdir


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
    tcol = 'tcol'
    start = 'start'
    end = 'end'
    cols = 'cols'

mc = MetaCol()

def list_dir(dirp: Path) :
    lo = list(dirp.glob('*'))
    lo = [i for i in lo if i.name != '.DS_Store']
    return lo

def update_meta(mp , dp) :
    with open(mp , 'r') as f :
        js = json.load(f)

    df = read_data(dp)

    if js[mc.tcol] is not None :
        if js[mc.tcol] in df.columns :
            js[mc.start] = df[js[mc.tcol]].min()
            js[mc.end] = df[js[mc.tcol]].max()

    js[mc.cols] = list(df.columns)

    with open(mp , 'w') as f :
        json.dump(js , f , indent = 4)

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
    _ = df.apply(lambda x : update_meta(x[c.mp] , x[c.dp]) , axis = 1)

    ##

##
if __name__ == "__main__" :
    main()
    print('Done!')
