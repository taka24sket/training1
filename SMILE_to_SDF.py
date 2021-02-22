# -*- coding: utf-8 -*-
"""
Created on Thu Jan 28 08:41:41 2021

@author: takashi.shiozawa
"""

import numpy as np
import pandas as pd
import os
from rdkit import Chem
from rdkit.Chem import AllChem
import sys

args = sys.argv

import_file = args[1]
export_file = args[2]

df = pd.read_excel(import_file)
#print(import_file)
#df.columns

names = df["ID"]
smiles = df["SMILES"]
mols = [Chem.MolFromSmiles(smile) for smile in smiles]
for mol in mols:
    if mol is None:
        continue
    #水素付加
    Chem.AddHs(mol)
    #2D構造
    AllChem.Compute2DCoords(mol)

writer = Chem.SDWriter(export_file)

for (mol,name)in zip(mols,names):
    if mol is None:
        continue
    mol.SetProp("name",name)
    writer.write(mol)

writer.close()
