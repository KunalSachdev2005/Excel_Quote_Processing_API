#**************************************************************************************************
# Program Name: Extraction.py
# Description: This program extracts required data from BSC Excel and uploads it into the Database.
# Created on 30-05-2022 by Kunal Sachdev
#**************************************************************************************************


def oricon_data_retriever_uploader(file_path):
    
    # Importing the required libraries
    import pandas as pd
    import os
    import sqlalchemy as db
    from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String
    import numpy as np

    # Creating connection to oracle database
    engine = create_engine('oracle://system:kunal@localhost:1521/orcl')
    con = engine.connect()
    meta = db.MetaData()

    # Reading Excel file in Python
    oricon_df = pd.read_excel(sheet_name = "Quotation Sheet", io = file_path, header = None)

    # Main DataFrame 1 and DataFrame 2
    df_1 = oricon_df.iloc[0:13, 0:2]
    df_2 = oricon_df.iloc[23:35, 0:8]
    lead_id = df_1[1][10]

    # Uploading values from the Main DataFrames to BSC_DA_TRANSACTION table in oracle

        # Setting isactive as 0
    transaction = db.Table("BSC_DA_TRANSACTION", meta, autoload = True, autoload_with = engine) 
    transaction_isactive_update_query = "update BSC_DA_TRANSACTION set BSC_DA_TRANSACTION.isactive = 0 where BSC_DA_TRANSACTION.trans_id = '"+str(lead_id)+"'"
    engine.execute(transaction_isactive_update_query)

        # Uploading the values
    transaction_insert_query = db.insert(transaction).values(trans_id = lead_id, industry = df_1[1][12], hazard_category = df_1[1][5], graded_retention = df_1[1][8], quality_score = df_1[1][6], nat_cat_score = df_1[1][7], isactive = 1, typeofrisk = df_2[2][33], username_var = df_2[2][24])
    engine.execute(transaction_insert_query)
    
    # Creating constants
    
    coverdetails = db.Table("BSC_DA_TRANS_COVERDETAILS", meta, autoload = True, autoload_with = engine)
    coverdetails_isactive_update_query = "update BSC_DA_TRANS_COVERDETAILS set BSC_DA_TRANS_COVERDETAILS.isactive = 0 where BSC_DA_TRANS_COVERDETAILS.trans_id = '"+str(lead_id)+"'"
    
    deductible_table = db.Table("BSC_DA_TRANS_DEDUCTIBLE", meta, autoload = True, autoload_with = engine)
    deductible_isactive_update_query = "update BSC_DA_TRANS_DEDUCTIBLE set BSC_DA_TRANS_DEDUCTIBLE.isactive = 0 where BSC_DA_TRANS_DEDUCTIBLE.trans_id = '"+str(lead_id)+"'"
    
    warranties_table = db.Table("BSC_DA_TRANS_WARRANTIES", meta, autoload = True, autoload_with = engine)
    warranties_isactive_update_query = "update BSC_DA_TRANS_WARRANTIES set BSC_DA_TRANS_WARRANTIES.isactive = 0 where BSC_DA_TRANS_WARRANTIES.trans_id = '"+str(lead_id)+"'"
    
    exclusion_table = db.Table("BSC_DA_TRANS_EXCLUSION", meta, autoload = True, autoload_with = engine)
    exclusion_isactive_update_query = "update BSC_DA_TRANS_EXCLUSION set BSC_DA_TRANS_EXCLUSION.isactive = 0 where BSC_DA_TRANS_EXCLUSION.trans_id = '"+str(lead_id)+"'"
    
    subjectivities_table = db.Table("BSC_DA_TRANS_SUBJECTIVITIES", meta, autoload = True, autoload_with = engine)
    subjectivities_isactive_update_query = "update BSC_DA_TRANS_SUBJECTIVITIES set BSC_DA_TRANS_SUBJECTIVITIES.isactive = 0 where BSC_DA_TRANS_SUBJECTIVITIES.trans_id = '"+str(lead_id)+"'"
    
    clauses_table = db.Table("BSC_DA_TRANS_CLAUSES", meta, autoload = True, autoload_with = engine)
    clauses_isactive_update_query = "update BSC_DA_TRANS_CLAUSES set BSC_DA_TRANS_CLAUSES.isactive = 0 where BSC_DA_TRANS_CLAUSES.trans_id = '"+str(lead_id)+"'"
    
    def row_getter(subrow, dataframe, rowid):
        for idx, rows in (dataframe.iterrows()):
            for ids, col in enumerate(dataframe.columns):
                if str(dataframe[col][idx]).lower().strip() == rowid:
                    subrow.append(idx+1)
    
    def reset_index(dataframe):
                dataframe = dataframe.reset_index()
                dataframe = dataframe.drop(columns = "index") if "index" in dataframe.columns else dataframe
                return dataframe
                    
    # Setting isactive as 0
    engine.execute(coverdetails_isactive_update_query)
    engine.execute(deductible_isactive_update_query)
    engine.execute(warranties_isactive_update_query)
    engine.execute(exclusion_isactive_update_query)
    engine.execute(subjectivities_isactive_update_query)
    engine.execute(clauses_isactive_update_query)
    
    
    
    def sfsp():
        """This function asks the user for the path of the Excel file, retrieves sfsp product's sub-sections and adds them to the appropriate tables in oracle."""
    
        # Creating sfsp_retriever function

        def sfsp_retriever():
            """This function retrieves the sfsp DataFrame from the oricon Excel File"""

            ins = ['standard fire & special perils policy', 'burglary insurance']

            user_req_ins = ['standard fire & special perils policy']

            ins_row = []
            for i in range(0, len(ins)):
                for a,b in (oricon_df.iterrows()):
                    for j,col in enumerate(oricon_df.columns):
                        if str(oricon_df[col][a]).lower().strip() == ins[i]:
                            ins_row.append(a)

            user_req_ins_row = []
            for i in range(0, len(user_req_ins)):
                for j in range(0, len(ins)):
                    if user_req_ins[i] == ins[j]:
                        user_req_ins_row.append(ins_row[j])

            for i in range(0, len(user_req_ins)):
                x = oricon_df.iloc[user_req_ins_row[i]:ins_row[ins.index(user_req_ins[i])+1], 0:20]
                x = x.reset_index()
                x = x.drop(columns = 'index') if 'index' in x.columns else x
                return(x)

        # Creating subsection_sfsp_retriever function

        def subsection_sfsp_retriever():
            """This function retrieves various sections of sfsp product from the oricon Excel File"""

            sfsp_df = sfsp_retriever()

            subsections = ['s. no', 'add-on covers', 'supplementary clauses and condition ( to be added below this part )', 'terms and conditions', "deductible -", "clauses", "supplementary clause and conditions - wording", "warranties", "exclusions", "subjectivities"]

            sub_row = []
            for i in subsections:
                for idx,rows in (sfsp_df.iterrows()):
                    for ids,col in enumerate(sfsp_df.columns):
                        if str(sfsp_df[col][idx]).lower().strip() == i:
                            sub_row.append(idx)
                            
            row_getter(sub_row, sfsp_df, '7.4')

            main_policies = sfsp_df.iloc[sub_row[0]:sub_row[1]]
            header = main_policies.iloc[0,:].values.tolist()
            main_policies.columns = header
            main_policies = main_policies.drop(1, axis = 0)
            add_on_policies = sfsp_df.iloc[(sub_row[1]+1):sub_row[2]]
            supplementary_policies = sfsp_df.iloc[(sub_row[2]+1):sub_row[3]]
            T_and_C = sfsp_df.iloc[(sub_row[3]+1):sub_row[4]]
            deductible = sfsp_df.iloc[(sub_row[4]+1):sub_row[5]]   
            clauses = sfsp_df.iloc[(sub_row[5]+1):sub_row[6]]
            supplementary_clauses_and_conditions = sfsp_df.iloc[(sub_row[6]+1):sub_row[7]]
            warranties = sfsp_df.iloc[(sub_row[7]+1):sub_row[8]]
            exclusions = sfsp_df.iloc[(sub_row[8]+1):sub_row[9]]
            subjectivities = sfsp_df.iloc[(sub_row[9]+1):sub_row[10]]

            add_on_policies.columns = header
            supplementary_policies.columns = header

            main_policies = reset_index(main_policies)
            add_on_policies = reset_index(add_on_policies)
            supplementary_policies = reset_index(supplementary_policies)
            T_and_C = reset_index(T_and_C)
            deductible = reset_index(deductible)
            clauses = reset_index(clauses)
            supplementary_clauses_and_conditions = reset_index(supplementary_clauses_and_conditions)
            warranties = reset_index(warranties)
            exclusions = reset_index(exclusions)
            subjectivities = reset_index(subjectivities)

            main_policies = main_policies.replace(np.nan, 0)
            add_on_policies = add_on_policies.replace(np.nan,0)
            supplementary_policies = supplementary_policies.replace(np.nan, 0)

            return main_policies, add_on_policies, supplementary_policies, T_and_C, deductible, clauses, supplementary_clauses_and_conditions, warranties, exclusions, subjectivities

        main_policies, add_on_policies, supplementary_policies, T_and_C, deductible, clauses, supplementary_clauses_and_conditions, warranties, exclusions, subjectivities = subsection_sfsp_retriever()

        # Converting the main_policies, add_on_policies, and supplementary_policies DataFrames to lists
        sfsp_main_policies = list(main_policies.itertuples(index = False, name = None))
        sfsp_add_on_policies = list(add_on_policies.itertuples(index = False, name = None))
        sfsp_supplementary_policies = list(supplementary_policies.itertuples(index = False, name = None))
        sfsp_deductible = list(deductible.itertuples(index = False, name = None))
        sfsp_clauses = list(clauses.itertuples(index = False, name = None))
        sfsp_supplementary_clauses_and_conditions = list(supplementary_clauses_and_conditions.itertuples(index = False, name = None))
        sfsp_warranties = list(warranties.itertuples(index = False, name = None))
        sfsp_exclusions = list(exclusions.itertuples(index = False, name = None))
        sfsp_subjectivities = list(subjectivities.itertuples(index = False, name = None))

        # Uploading the values from above first 3 lists to BSC_DA_TRANS_COVERDETAILS table in oracle database

            # Uploading the values
        for i in range(len(sfsp_main_policies)):
            query1 = db.insert(coverdetails).values(coverlist_id = 'sfsp', isbenefit = sfsp_main_policies[i][2], benefit_desc = sfsp_main_policies[i][1], col_value = sfsp_main_policies[i][1], column_name = 'main_policies', trans_id = lead_id, isactive = 1)
            engine.execute(query1)
        for i in range(len(sfsp_add_on_policies)):
            query2 = db.insert(coverdetails).values(coverlist_id = 'sfsp', isbenefit = sfsp_add_on_policies[i][2], benefit_desc = sfsp_add_on_policies[i][1], col_value = sfsp_add_on_policies[i][1], column_name = 'add_on_policies', trans_id = lead_id, isactive = 1)
            engine.execute(query2)
        for i in range(len(sfsp_supplementary_policies)):
            query3 = db.insert(coverdetails).values(coverlist_id = 'sfsp', isbenefit = sfsp_supplementary_policies[i][2], benefit_desc = sfsp_supplementary_policies[i][1], col_value = sfsp_supplementary_policies[i][1], column_name = 'supplementary_policies', trans_id = lead_id, isactive = 1)
            engine.execute(query3)

        # Uploading the values from sfsp_deductible to BSC_DA_TRANS_DEDUCTIBLE table in oracle database

             # Uploading the values
        for i in range(len(sfsp_deductible)):
            query4 = db.insert(deductible_table).values(cover_id = 'sfsp', trans_id = lead_id, isactive = 1, deductible = str(sfsp_deductible[i][1]))
            engine.execute(query4)

         # Uploading the values from sfsp_warranties to BSC_DA_TRANS_WARRANTIES table in oracle database

            # Uploading the values
        for i in range(len(sfsp_warranties)):
            query5 = db.insert(warranties_table).values(cover_id = 'sfsp', trans_id = lead_id, isactive = 1, warranties = str(sfsp_warranties[i][1]))
            engine.execute(query5)

        # Uploading the values from sfsp_exclusions to BSC_DA_TRANS_EXCLUSION table in oracle database

            # Uploading the values
        for i in range(len(sfsp_exclusions)):
            query6 = db.insert(exclusion_table).values(cover_id = 'sfsp', trans_id = lead_id, isactive = 1, exclusion_desc = str(sfsp_exclusions[i][1]))
            engine.execute(query6)

        # Uploading the values from sfsp_subjectivities to BSC_DA_TRANS_SUBJECTIVITIES table in oracle database

            # Uploading the values
        for i in range(len(sfsp_subjectivities)):
            query7 = db.insert(subjectivities_table).values(cover_id = 'sfsp', trans_id = lead_id, isactive = 1, subjectives = str(sfsp_subjectivities[i][1]))
            engine.execute(query7)

        # Uploading the values from clauses and supplementary_clauses lists to BSC_DA_TRANS_CLAUSES table in oracle database

            # Uploading the values
        for i in range(len(sfsp_clauses)):
            query8 = db.insert(clauses_table).values(cover_id = 'sfsp', clause_desc = str(sfsp_clauses[i][1]), clause_id = 'clauses', trans_id = lead_id, isactive = 1)
            engine.execute(query8)
        for i in range(len(sfsp_supplementary_clauses_and_conditions)):
            query9 = db.insert(clauses_table).values(cover_id = 'sfsp', clause_desc = str(sfsp_supplementary_clauses_and_conditions[i][1]), clause_id = 'supplementary_clauses_and_conditions', trans_id = lead_id, isactive = 1)
            engine.execute(query9)
            
            
            
    def burglary():
        """This function asks the user for the path of the Excel file, retrieves burglary product's sub-sections and adds them to the appropriate tables in oracle."""
        
        # Creating burglary_retriever function
        
        def burglary_retriever():
            """This function retrieves the burglary DataFrame from the oricon Excel File"""
    
            ins = ['burglary insurance', 'plate glass insurance']
    
            user_req_ins = ['burglary insurance']
    
            ins_row = []
            for i in range(0, len(ins)):
                for a,b in (oricon_df.iterrows()):
                    for j,col in enumerate(oricon_df.columns):
                        if str(oricon_df[col][a]).lower().strip() == ins[i]:
                            ins_row.append(a)
    
            user_req_ins_row = []
            for i in range(0, len(user_req_ins)):
                for j in range(0, len(ins)):
                    if user_req_ins[i] == ins[j]:
                        user_req_ins_row.append(ins_row[j])
    
            for i in range(0, len(user_req_ins)):
                x = oricon_df.iloc[user_req_ins_row[i]:ins_row[ins.index(user_req_ins[i])+1], 0:20]
                x = x.reset_index()
                x = x.drop(columns = 'index') if 'index' in x.columns else x
                return(x)
            
        # Creating subsection_burglary_retriever function
        
        def subsection_burglary_retriever():
            """This function retrieves various sections of sfsp product from the oricon Excel File"""
    
            burglary_df = burglary_retriever()
    
            subsections = ['s. no', 'add-on covers', 'supplementary clauses and condition ( to be added below this part )', 'terms and conditions', "deductible -", "clauses", "supplementary clause and conditions - wording ( to be suitably modified by uw in case of bound risk )", "warranties", "exclusions", "subjectivities"]
    
            sub_row = []
            for i in subsections:
                for idx,rows in (burglary_df.iterrows()):
                    for ids,col in enumerate(burglary_df.columns):
                        if str(burglary_df[col][idx]).lower().strip() == i:
                            sub_row.append(idx)
                            
            row_getter(sub_row, burglary_df, '6.3')
            
            main_policies = burglary_df.iloc[sub_row[0]:sub_row[1]]
            header = main_policies.iloc[0,:].values.tolist()
            main_policies.columns = header
            main_policies = main_policies.drop(1, axis = 0)
            add_on_policies = burglary_df.iloc[(sub_row[1]+1):sub_row[2]]
            supplementary_policies = burglary_df.iloc[(sub_row[2]+1):sub_row[3]]
            T_and_C = burglary_df.iloc[(sub_row[3]+1):sub_row[4]]
            deductible = burglary_df.iloc[(sub_row[4]+1):sub_row[5]]   
            clauses = burglary_df.iloc[(sub_row[5]+1):sub_row[6]]
            supplementary_clauses_and_conditions = burglary_df.iloc[(sub_row[6]+1):sub_row[7]]
            warranties = burglary_df.iloc[(sub_row[7]+1):sub_row[8]]
            exclusions = burglary_df.iloc[(sub_row[8]+1):sub_row[9]]
            subjectivities = burglary_df.iloc[(sub_row[9]+1):sub_row[10]]
    
            add_on_policies.columns = header
            supplementary_policies.columns = header
    
            main_policies = reset_index(main_policies)
            add_on_policies = reset_index(add_on_policies)
            supplementary_policies = reset_index(supplementary_policies)
            T_and_C = reset_index(T_and_C)
            deductible = reset_index(deductible)
            clauses = reset_index(clauses)
            supplementary_clauses_and_conditions = reset_index(supplementary_clauses_and_conditions)
            warranties = reset_index(warranties)
            exclusions = reset_index(exclusions)
            subjectivities = reset_index(subjectivities)
            
            main_policies = main_policies.replace(np.nan, 0)
            add_on_policies = add_on_policies.replace(np.nan,0)
            add_on_policies = add_on_policies.replace("YES", 1)
            add_on_policies = add_on_policies.replace("NO", 0)
            supplementary_policies = supplementary_policies.replace(np.nan, 0)
    
            return main_policies, add_on_policies, supplementary_policies, T_and_C, deductible, clauses, supplementary_clauses_and_conditions, warranties, exclusions, subjectivities
        
        main_policies, add_on_policies, supplementary_policies, T_and_C, deductible, clauses, supplementary_clauses_and_conditions, warranties, exclusions, subjectivities = subsection_burglary_retriever()
        
        # Converting the main_policies, add_on_policies, and supplementary_policies DataFrames to lists
        burglary_main_policies = list(main_policies.itertuples(index = False, name = None))
        burglary_add_on_policies = list(add_on_policies.itertuples(index = False, name = None))
        burglary_supplementary_policies = list(supplementary_policies.itertuples(index = False, name = None))
        burglary_deductible = list(deductible.itertuples(index = False, name = None))
        burglary_clauses = list(clauses.itertuples(index = False, name = None))
        burglary_supplementary_clauses_and_conditions = list(supplementary_clauses_and_conditions.itertuples(index = False, name = None))
        burglary_warranties = list(warranties.itertuples(index = False, name = None))
        burglary_exclusions = list(exclusions.itertuples(index = False, name = None))
        burglary_subjectivities = list(subjectivities.itertuples(index = False, name = None))
        
        # Uploading the values from above first 3 lists to BSC_DA_TRANS_COVERDETAILS table in oracle database
        
            # Uploading the values
        for i in range(len(burglary_main_policies)):
            query1 = db.insert(coverdetails).values(coverlist_id = 'burglary', isbenefit = str(burglary_main_policies[i][2]), benefit_desc = burglary_main_policies[i][1], col_value = burglary_main_policies[i][1], column_name = 'main_policies', trans_id = lead_id, isactive = 1)
            engine.execute(query1)
        for i in range(len(burglary_add_on_policies)):
            query2 = db.insert(coverdetails).values(coverlist_id = 'burglary', isbenefit = str(burglary_add_on_policies[i][2]), benefit_desc = burglary_add_on_policies[i][1], col_value = burglary_add_on_policies[i][1], column_name = 'add_on_policies', trans_id = lead_id, isactive = 1)
            engine.execute(query2)
        for i in range(len(burglary_supplementary_policies)):
            query3 = db.insert(coverdetails).values(coverlist_id = 'burglary', isbenefit = str(burglary_supplementary_policies[i][2]), benefit_desc = burglary_supplementary_policies[i][1], col_value = burglary_supplementary_policies[i][1], column_name = 'supplementary_policies', trans_id = lead_id, isactive = 1)
            engine.execute(query3)
            
        # Uploading the values from burglary_deductible to BSC_DA_TRANS_DEDUCTIBLE table in oracle database
        
            # Uploading the values
        for i in range(len(burglary_deductible)):
            query4 = db.insert(deductible_table).values(cover_id = 'burglary', trans_id = lead_id, isactive = 1, deductible = str(burglary_deductible[i][1]))
            engine.execute(query4)
        
         # Uploading the values from burglary_warranties to BSC_DA_TRANS_WARRANTIES table in oracle database
            
            # Uploading the values
        for i in range(len(burglary_warranties)):
            query5 = db.insert(warranties_table).values(cover_id = 'burglary', trans_id = lead_id, isactive = 1, warranties = str(burglary_warranties[i][1]))
            engine.execute(query5)
            
        # Uploading the values from burglary_exclusions to BSC_DA_TRANS_EXCLUSION table in oracle database
        
            # Uploading the values
        for i in range(len(burglary_exclusions)):
            query6 = db.insert(exclusion_table).values(cover_id = 'burglary', trans_id = lead_id, isactive = 1, exclusion_desc = str(burglary_exclusions[i][1]))
            engine.execute(query6)
            
        # Uploading the values from burglary_subjectivities to BSC_DA_TRANS_SUBJECTIVITIES table in oracle database
        
            # Uploading the values
        for i in range(len(burglary_subjectivities)):
            query7 = db.insert(subjectivities_table).values(cover_id = 'burglary', trans_id = lead_id, isactive = 1, subjectives = str(burglary_subjectivities[i][1]))
            engine.execute(query7)
            
        # Uploading the values from clauses and supplementary_clauses lists to BSC_DA_TRANS_CLAUSES table in oracle database
            
            # Uploading the values
        for i in range(len(burglary_clauses)):
            query8 = db.insert(clauses_table).values(cover_id = 'burglary', clause_desc = str(burglary_clauses[i][1]), clause_id = 'clauses', trans_id = lead_id, isactive = 1)
            engine.execute(query8)
        for i in range(len(burglary_supplementary_clauses_and_conditions)):
            query9 = db.insert(clauses_table).values(cover_id = 'burglary', clause_desc = str(burglary_supplementary_clauses_and_conditions[i][1]), clause_id = 'supplementary_clauses_and_conditions', trans_id = lead_id, isactive = 1)
            engine.execute(query9)
            
            
            
    def plate_glass():
        """This function asks the user for the path of the Excel file, retrieves plate glass product's sub-sections and adds them to the appropriate tables in oracle."""
    
        # Creating plate_glass_retriever function
        
        def plate_glass_retriever():
            """This function retrieves the plate glass DataFrame from the oricon Excel File"""
    
            ins = ['plate glass insurance', 'neon sign/ glow sign all risk insurance']
    
            user_req_ins = ['plate glass insurance']
    
            ins_row = []
            for i in range(0, len(ins)):
                for a,b in (oricon_df.iterrows()):
                    for j,col in enumerate(oricon_df.columns):
                        if str(oricon_df[col][a]).lower().strip() == ins[i]:
                            ins_row.append(a)
    
            user_req_ins_row = []
            for i in range(0, len(user_req_ins)):
                for j in range(0, len(ins)):
                    if user_req_ins[i] == ins[j]:
                        user_req_ins_row.append(ins_row[j])
    
            for i in range(0, len(user_req_ins)):
                x = oricon_df.iloc[user_req_ins_row[i]:ins_row[ins.index(user_req_ins[i])+1], 0:20]
                x = x.reset_index()
                x = x.drop(columns = 'index') if 'index' in x.columns else x
                return(x)
            
        # Creating subsection_plate_glass_retriever function
        
        def subsection_plate_glass_retriever():
            """This function retrieves various sections of sfsp product from the oricon Excel File"""
    
            plate_glass_df = plate_glass_retriever()
    
            #subsections = ['s. no', 'add-on covers', 'supplementary clauses and condition ( to be added below this part )', 'terms and conditions', "deductible -", "clauses", "supplementary clause and conditions - wording ( to be suitably modified by uw in case of bound risk )", "warranties", "exclusions", "subjectivities"]
            subsections = ['s. no', 'supplementary clauses and condition ( to be added below this part )', 'terms and conditions', "deductible -", "clauses", "supplementary clause and conditions - wording ( to be suitably modified by uw in case of bound risk )", "warranties", "exclusions", "subjectivities"]
            
            sub_row = []
            for i in subsections:
                for idx,rows in (plate_glass_df.iterrows()):
                    for ids,col in enumerate(plate_glass_df.columns):
                        if str(plate_glass_df[col][idx]).lower().strip() == i:
                            sub_row.append(idx)
                            
            row_getter(sub_row, plate_glass_df, '6.2')
            
            main_policies = plate_glass_df.iloc[sub_row[0]:sub_row[1]]
            header = main_policies.iloc[0,:].values.tolist()
            main_policies.columns = header
            main_policies = main_policies.drop(1, axis = 0)
            #add_on_policies = df.iloc[(sub_row[1]+1):sub_row[2]]
            supplementary_policies = plate_glass_df.iloc[(sub_row[1]+1):sub_row[2]]
            T_and_C = plate_glass_df.iloc[(sub_row[2]+1):sub_row[3]]
            deductible = plate_glass_df.iloc[(sub_row[3]+1):sub_row[4]]   
            clauses = plate_glass_df.iloc[(sub_row[4]+1):sub_row[5]]
            supplementary_clauses_and_conditions = plate_glass_df.iloc[(sub_row[5]+1):sub_row[6]]
            warranties = plate_glass_df.iloc[(sub_row[6]+1):sub_row[7]]
            exclusions = plate_glass_df.iloc[(sub_row[7]+1):sub_row[8]]
            subjectivities = plate_glass_df.iloc[(sub_row[8]+1):sub_row[9]]
    
            #add_on_policies.columns = header
            supplementary_policies.columns = header
    
            main_policies = reset_index(main_policies)
            #add_on_policies = reset_index(add_on_policies)
            supplementary_policies = reset_index(supplementary_policies)
            T_and_C = reset_index(T_and_C)
            deductible = reset_index(deductible)
            clauses = reset_index(clauses)
            supplementary_clauses_and_conditions = reset_index(supplementary_clauses_and_conditions)
            warranties = reset_index(warranties)
            exclusions = reset_index(exclusions)
            subjectivities = reset_index(subjectivities)
            
            main_policies = main_policies.replace(np.nan, 0)
            #add_on_policies = add_on_policies.replace(np.nan,0)
            supplementary_policies = supplementary_policies.replace(np.nan, 0)
    
            #return main_policies, add_on_policies, supplementary_policies, T_and_C, deductible, clauses, supplementary_clauses_and_conditions, warranties, exclusions, subjectivities
            return main_policies, supplementary_policies, T_and_C, deductible, clauses, supplementary_clauses_and_conditions, warranties, exclusions, subjectivities
        
        #main_policies, add_on_policies, supplementary_policies, T_and_C, deductible, clauses, supplementary_clauses_and_conditions, warranties, exclusions, subjectivities = subsection_plate_glass_retriever()
        main_policies, supplementary_policies, T_and_C, deductible, clauses, supplementary_clauses_and_conditions, warranties, exclusions, subjectivities = subsection_plate_glass_retriever()
        
        # Converting the main_policies, add_on_policies, and supplementary_policies DataFrames to lists
        plate_glass_main_policies = list(main_policies.itertuples(index = False, name = None))
        #plate_glass_add_on_policies = list(add_on_policies.itertuples(index = False, name = None))
        plate_glass_supplementary_policies = list(supplementary_policies.itertuples(index = False, name = None))
        plate_glass_deductible = list(deductible.itertuples(index = False, name = None))
        plate_glass_clauses = list(clauses.itertuples(index = False, name = None))
        plate_glass_supplementary_clauses_and_conditions = list(supplementary_clauses_and_conditions.itertuples(index = False, name = None))
        plate_glass_warranties = list(warranties.itertuples(index = False, name = None))
        plate_glass_exclusions = list(exclusions.itertuples(index = False, name = None))
        plate_glass_subjectivities = list(subjectivities.itertuples(index = False, name = None))
        
        # Uploading the values from above first 2 lists to BSC_DA_TRANS_COVERDETAILS table in oracle database
        
            # Uploading the values
        for i in range(len(plate_glass_main_policies)):
            query1 = db.insert(coverdetails).values(coverlist_id = 'plate_glass', isbenefit = str(plate_glass_main_policies[i][2]), benefit_desc = plate_glass_main_policies[i][1], col_value = plate_glass_main_policies[i][1], column_name = 'main_policies', trans_id = lead_id, isactive = 1)
            engine.execute(query1)
        #for i in range(len(plate_glass_add_on_policies)):
            #query2 = db.insert(coverdetails).values(coverlist_id = 'plate_glass', isbenefit = str(plate_glass_add_on_policies[i][2]), benefit_desc = plate_glass_add_on_policies[i][1], col_value = plate_glass_add_on_policies[i][1], column_name = 'add_on_policies', trans_id = lead_id, isactive = 1)
            #engine.execute(query2)
        for i in range(len(plate_glass_supplementary_policies)):
            query3 = db.insert(coverdetails).values(coverlist_id = 'plate_glass', isbenefit = str(plate_glass_supplementary_policies[i][2]), benefit_desc = plate_glass_supplementary_policies[i][1], col_value = plate_glass_supplementary_policies[i][1], column_name = 'supplementary_policies', trans_id = lead_id, isactive = 1)
            engine.execute(query3)
            
        # Uploading the values from plate_glass_deductible to BSC_DA_TRANS_DEDUCTIBLE table in oracle database
        
            # Uploading the values
        for i in range(len(plate_glass_deductible)):
            query4 = db.insert(deductible_table).values(cover_id = 'plate_glass', trans_id = lead_id, isactive = 1, deductible = str(plate_glass_deductible[i][1]))
            engine.execute(query4)
        
         # Uploading the values from plate_glass_warranties to BSC_DA_TRANS_WARRANTIES table in oracle database
            
            # Uploading the values
        for i in range(len(plate_glass_warranties)):
            query5 = db.insert(warranties_table).values(cover_id = 'plate_glass', trans_id = lead_id, isactive = 1, warranties = str(plate_glass_warranties[i][1]))
            engine.execute(query5)
            
        # Uploading the values from plate_glass_exclusions to BSC_DA_TRANS_EXCLUSION table in oracle database
        
            # Uploading the values
        for i in range(len(plate_glass_exclusions)):
            query6 = db.insert(exclusion_table).values(cover_id = 'plate_glass', trans_id = lead_id, isactive = 1, exclusion_desc = str(plate_glass_exclusions[i][1]))
            engine.execute(query6)
            
        # Uploading the values from plate_glass_subjectivities to BSC_DA_TRANS_SUBJECTIVITIES table in oracle database
        
            # Uploading the values
        for i in range(len(plate_glass_subjectivities)):
            query7 = db.insert(subjectivities_table).values(cover_id = 'plate_glass', trans_id = lead_id, isactive = 1, subjectives = str(plate_glass_subjectivities[i][1]))
            engine.execute(query7)
            
        # Uploading the values from clauses and supplementary_clauses lists to BSC_DA_TRANS_CLAUSES table in oracle database
        
            # Uploading the values
        for i in range(len(plate_glass_clauses)):
            query8 = db.insert(clauses_table).values(cover_id = 'plate_glass', clause_desc = str(plate_glass_clauses[i][1]), clause_id = 'clauses', trans_id = lead_id, isactive = 1)
            engine.execute(query8)
        for i in range(len(plate_glass_supplementary_clauses_and_conditions)):
            query9 = db.insert(clauses_table).values(cover_id = 'plate_glass', clause_desc = str(plate_glass_supplementary_clauses_and_conditions[i][1]), clause_id = 'supplementary_clauses_and_conditions', trans_id = lead_id, isactive = 1)
            engine.execute(query9)        
    
    
    sfsp()
    burglary()
    plate_glass()