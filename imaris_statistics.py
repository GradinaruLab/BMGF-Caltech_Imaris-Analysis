# -*- coding: utf-8 -*-
"""
Created on Fri Dec 10 15:34:21 2021

@author: Carrie
"""

import pandas as pd

test_sheet = "AAV7 GFP Surface.xls"
actin_test = "AAV7 Actin Surface.xls"
gfp_test = "AAV7 GFP Surface.xls"

test_prefixes = ["AAV2", "AAV7"]


def getVolSheet(excel_name, sheet = "Volume"):
    """
    Reads an excel file of data exported from Imaris and returns a data frame
    containing the volume of all objects in the surface, and the image where
    that surface was generated. 
    
    Currently hardcoded for sheets with columns in a specific order. 
    

    Parameters
    ----------
    excel_name : str
        The name of the excel file from which to read data. 
    sheet : str, optional
        Which sheet of the excel file to read. The default is "Volume".

    Returns
    -------
    out_df : pandas.dataFrame
        a data frame with two columns: the volume of each object in the surface,
        and the image where the surface was generated. 

    """
    df = pd.read_excel(excel_name, sheet_name = sheet, header=[0,1])
    vol_col = df.iloc[:,:1]
    img_col = df.iloc[:,9:10]
    
    out_df = vol_col.merge(img_col, how='outer', left_index=True, right_index=True)
    #vol_data.rename(index={1:vol_name})
    #source_img_data.rename("Unnamed : 9":"cat")
    
    #return source_img_data
    return out_df

def sumVolumes(df):
    """
    Sums all volumes for each image name and returns a dictionary of the 
    summed volumes keyed to the names. 
    
    Harcoded to expect column names "Volume" and "Original Image Name"

    Parameters
    ----------
    df : pandas.dataFrame
        A data frame from getVolSheet. Expects column names "Volume" and 
        "Original Image Name"

    Returns
    -------
    vol_dict : dict of str:float
        A dictionary where the keys are image names and the values are 
        the sum of the volume of the surface. 

    """
    images = df["Volume"]["Original Image Name"].unique()
    vol_dict = {}
    for img in images:
        vol_dict[img] = 0
    for i in range(len(df)):
        img_i = df["Volume"]["Original Image Name"][i]
        vol_i = df["Volume"]["Volume"][i]
        vol_dict[img_i] += vol_i
    return vol_dict
        
def pairActinGFP(actin_sheet, GFP_sheet):
    """
    Generates summed volumes for an actin excel sheet and a gfp excel sheet,
    and creates a data frame with the actin and gfp volume for each image

    Parameters
    ----------
    actin_sheet : TYPE
        DESCRIPTION.
    GFP_sheet : TYPE
        DESCRIPTION.

    Returns
    -------
    df : TYPE
        DESCRIPTION.

    """
    A_df = getVolSheet(actin_sheet)
    A_volumes = sumVolumes(A_df) #image name : summed actin volume
    G_df = getVolSheet(GFP_sheet) 
    G_volumes = sumVolumes(G_df) #image name : summed GFP volume
    data_dict = {"Original Image Name" : [], 
                 "Actin Volume" : [], "GFP Volume" : []}
    for key in A_volumes.keys():
        
        data_dict["Original Image Name"].append(key)
        data_dict["Actin Volume"].append(A_volumes[key])
        if key in G_volumes.keys():
            data_dict["GFP Volume"].append(G_volumes[key])
        else: #in case there's no surface at all
            data_dict["GFP Volume"].append(0)
    
    df = pd.DataFrame.from_dict(data_dict)
    return df

def generateAGFileNames(prefix_list):
    """
    Generates a list of tuples of (actin,gfp) file names from which we want
    to get the volumes. The file names will be prefix + " Actin Surface.xls" 
    and prefix + " GFP Surface.xls"

    Parameters
    ----------
    prefix_list : list of str
        The prefixes of file names to be processed.
        (Suffix of filename should indicate actin or gfp)
        NOTE: these must be files that already exist in folder

    Returns
    -------
    pair_list : list of tuples of strings [(str,str)]
        A list of tuples of filenames ordered ("actin", "gfp")

    """
    pair_list = []
    for p in prefix_list:
        actin = p + " Actin Surface.xls"
        gfp = p + " GFP Surface.xls"
        pair_list.append((actin,gfp))
    return pair_list

def volumesToCSV(prefix_list, out_csv):
    """
    Gets and sums the volumes of actin and gfp surfaces from excel sheets 
    exported from Imaris. Generates a csv file with the actin and gfp volume
    for each image.

    Parameters
    ----------
    prefix_list : list of str
        The prefixes of file names to be processed.
        (Suffix of filename should indicate actin or gfp)
        NOTE: these must be files that already exist in folder
    out_csv : str
        the name of the csv file to write to. Will be overwritten if it 
        already exists. 

    Returns
    -------
    None.

    """
    file_pairs = generateAGFileNames(prefix_list)
    for i in range(len(file_pairs)):
        df = pairActinGFP(file_pairs[i][0], file_pairs[i][1])
        condition = [prefix_list[i]] * len(df)
        df["Condition"] = condition
        if i == 0:
            df.to_csv(out_csv, mode='w', index=False, header=True)
        else:
            df.to_csv(out_csv, mode='a', index=False, header=False)
    
    
    
    
####    commands to generate output CSVs
sero_prefixes = ["ni-sero", "AAV2", "AAV6", "AAV7", "AAV8", "AAVDJ"]
conc_prefixes = ["ni-conc", "low", "medium", "high", "extra high", "XXH"]
volumesToCSV(sero_prefixes, "serotype_volumes.csv")
volumesToCSV(conc_prefixes, "concentration_volumes.csv")
    
    

#df = getVolSheet(test_sheet)
#d = sumVolumes(df) 
#p = pairActinGFP(actin_test,gfp_test)
#l = generateAGFileNames(sero_prefixes)
#volumesToCSV(test_prefixes, "test_out.csv")





#Volume	Unit	Category	Birth [s]	Death [s]	ID	OriginalID	Original 
#Component Name	Original Component ID	Original Image Name	Original Image ID

