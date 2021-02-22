# -*- coding: utf-8 -*-
"""
Created on Wed Feb  3 18:48:53 2021

@author: takashi.shiozawa
"""

import pandas as pd

from rdkit import Chem
from rdkit.Chem import Draw
from rdkit.Chem import Descriptors
from rdkit.ML.Descriptors import MoleculeDescriptors
from rdkit.Chem import PandasTools

# 適当な化学構造のsmilesリストを準備する
smiles_list = ['Oc(cccc1)c1O', 'OC(C(C1)C1(Br)Br)=O', 'Cc(cc1)ccc1O',
       'Oc(cc1)ccc1Cl', 'OC(c1cocc1)=O', 'CC(C)c1nnn[nH]1',
       'CN(C=C1)C=CC1=N', 'CCc1cccnn1', 'C[C@H]([C@H]1NC)[C@@H]1NC',
       'CCCc1ncc[nH]1']

# 化合物のラベルを作成
label_list = ['sample_{}'.format(i) for i in range(len(smiles_list))]

# molオブジェクトのリストを作成
mols_list = [Chem.MolFromSmiles(smile) for smile in smiles_list]


# RDkit記述子の作成
descriptor_names = [descriptor_name[0] for descriptor_name in Descriptors._descList[:5]]
descriptor_calculation = MoleculeDescriptors.MolecularDescriptorCalculator(descriptor_names)
RDkit = [descriptor_calculation.CalcDescriptors(mol_temp) for mol_temp in mols_list]

df = pd.DataFrame(RDkit, columns = descriptor_names,index=label_list)
df['smiles'] = smiles_list

# DataFrameへのImageの追加とエクセルファイルでの出力
PandasTools.AddMoleculeColumnToFrame(df, molCol='IMAGE', smilesCol='smiles')
PandasTools.SaveXlsxFromFrame(df, 'data_frame.xlsx',molCol='IMAGE',size=(150,150))
