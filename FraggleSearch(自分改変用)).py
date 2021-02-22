from rdkit import rdBase, Chem, DataStructs
from rdkit.Chem import Draw, AllChem, PandasTools, rdMolDescriptors#, MolStandardize
from rdkit.Chem.Draw import rdMolDraw2D
from rdkit.Chem.Fraggle import FraggleSim
import pandas as pd
import csv
import sys

# リファレンス読み込み
args = sys.argv #pythonスクリプトに渡されたコマンドラインの引数
print("Reference File:", args[1])　#ファイルの表示
filename = args[1]
suppl = Chem.SDMolSupplier("Namiki_202006_random_1000.sdf", sanitize=False)
reference = [x for x in suppl if x is not None] #SDMolSupplierは入力不備があるとNoneを返すため

# クエリ読み込み
print("Query File:", args[2])
suppl2 = Chem.SDMolSupplier("FCP-675.sdf")
query = [x for x in suppl2 if x is not None] #上記と同様

# 前処理
reference_desalt = []
count=0
for mol in reference:
    # 無効構造チェック
    if mol is None:
        print("skip") #Noneのものがあればskipを表示すう
        continue 
    # サニタイズチェック
    try:
        Chem.SanitizeMol(mol)
    except ValueError as e:
        print("サニタイズエラー:",e) #try「実行したい処理」except例外発生時の処理
        continue
    # 標準化
    normalizer = MolStandardize.normalize.Normalizer()
    test_mol_norm = normalizer.normalize(mol)
    # 脱塩(一番大きいフラグメントを残す)
    lfc = MolStandardize.fragment.LargestFragmentChooser()
    test_mol_desalt = lfc.choose(test_mol_norm)
    # 電荷の中和
    uc = MolStandardize.charge.Uncharger()
    test_mol_neu = uc.uncharge(test_mol_desalt)
    # test_mol_neu.SetProp("NScode", mol.GetProp("NScode"))
    reference_desalt.append(test_mol_neu)
    count = count+1

print("Reference Molecule Num: ", len(reference_desalt))

# Morganによる類似性評価
morgan_reference = [AllChem.GetMorganFingerprintAsBitVect(mol, 2, 2048) for mol in reference_desalt] #(mol,中心からの半径,bit数)
morgan_query = [AllChem.GetMorganFingerprintAsBitVect(mol, 2, 2048) for mol in query]
tanimoto = DataStructs.BulkTanimotoSimilarity(morgan_query[0], morgan_reference)


# Fraggleによる類似性評価
fraggle_sim = []
fraggle_match = []
for (sim, match) in [FraggleSim.GetFraggleSimilarity(query[0], reference_desalt[i]) for i in range(len(reference_desalt))] :
    fraggle_sim.append(sim)
    fraggle_match.append(match)


# 結果をPDFファイルに書き込み
# img = Draw.MolsToGridImage(namiki[:16], molsPerRow=4, subImgSize=(300,200), legends=['Fraggle: {:.2f}'.format(i) for i in fraggle_similarity[:16]])
# img.save('./fraggle.pdf')

# 結果をSDFファイルに書き込み
sdf_filename = "./fraggle-result.sdf"
writer = Chem.SDWriter(sdf_filename)
query[0].SetProp("NScode", "2-17") #molオブジェクトにプロパティを付加する
query[0].SetProp("Tanimoto_Sim", "1")
query[0].SetProp("Fraggle_Sim", "1")
writer.write(query[0])
for(mol_desalt, tani, fra, match) in zip(reference_desalt, tanimoto, fraggle_sim, fraggle_match):
    mol_desalt.SetProp("Tanimoto_Sim", str(tani))
    mol_desalt.SetProp("Fraggle_Sim", str(fra))
    mol_desalt.SetProp("Fraggle_Match", str(match))
    try:
        writer.write(mol_desalt)
    except ValueError as e:
        continue
writer.close()

# 結果をTSVに書き込み
tsv_filename = "./fraggle-result.tsv"
with open(tsv_filename, 'w') as f: #"w"は書き込みモード
    writer = csv.writer(f, delimiter='\t') #delimiterは区切りを決めるこの場合tab区切りを指定
    l = ["smiles", "tanimoto", "fraggle"]
    writer.writerow(l) #1行ごとに書き込む
    for(mol_desalt, tani, fra) in zip(reference_desalt, tanimoto, fraggle_sim):
        l = [Chem.MolToSmiles(mol_desalt), tani, fra]
        writer.writerow(l)

print("Result File(SDF) : ", sdf_filename)
print("Result File(TSV) : ", tsv_filename)