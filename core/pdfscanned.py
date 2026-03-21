# Copyright (c) 2024 Bhanusri mattey
# Licensed under the Business Source License 1.1
# See LICENSE file in the project root for details.
# Commercial use prohibited until 2030-03-01

from core.pdfhelper import *   # TextDetection, EdgeDetection, remove_page_border_edges, ExcelWriter, etc.
from core.deskew import * 
import json  
from core.pdftext import*
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent.parent  # project_root


pdf_path = "office10.pdf"
on_page_done = None

maxTolerances = 70
def TableDetectionComp(h_edges,v_edges,tableData,tableList,fixed,indexList):
     """Returns a valid table running through all tolerances if no table found it tries to
     forms table with the available edges"""
     with open(BASE_DIR / "tolerances" / "tolerances_3.json", "r") as f:
        tolerance_configs_table = json.load(f)
     indexList2 = {}
     index = 0
     while len(h_edges) != 0 and len(v_edges) != 0:
            
            
        if not indexList or (index in indexList and indexList[index] != maxTolerances):   
            found,tableD,table,h_edges,v_edges,isValid,idx = TableDetection(fixed,h_edges,v_edges,indexList,index)
        
        if idx < maxTolerances:
            indexList2[index] = idx+1
        
        elif idx == maxTolerances:
            indexList2[index] = idx

        if tableD != None and isValid:
            tableData.append(tableD)
                
                
        if table != None and isValid:
            tableList.append(table)
        
        
        if table == None and isValid:
            break
        
        
        if found:
            
            
            
            for k in range(len(tolerance_configs_table)):  
                cfg_2 = tolerance_configs_table[k]
                tsx   = cfg_2["TABLE_SNAP_X_TOLERANCE"]
                tsy  = cfg_2["TABLE_SNAP_Y_TOLERANCE"]
                tjx  = cfg_2["TABLE_JOIN_X_TOLERANCE"]
                tjy   = cfg_2["TABLE_JOIN_Y_TOLERANCE"]
                esx  = cfg_2["EDGE_SNAP_X_TOLERANCE"]
                esy  = cfg_2["EDGE_SNAP_Y_TOLERANCE"]
                ejx  = cfg_2["EDGE_JOIN_X_TOLERANCE"]
                ejy  = cfg_2["EDGE_JOIN_Y_TOLERANCE"]
                
                            
                tableP.TABLE_SNAP_X_TOLERANCE = tsx
                tableP.TABLE_SNAP_Y_TOLERANCE = tsy
                tableP.TABLE_JOIN_X_TOLERANCE = tjx
                tableP.TABLE_JOIN_Y_TOLERANCE = tjy
                tableP.EDGE_SNAP_X_TOLERANCE = esx
                tableP.EDGE_SNAP_Y_TOLERANCE = esy
                tableP.EDGE_JOIN_X_TOLERANCE = ejx
                tableP.EDGE_JOIN_Y_TOLERANCE = ejy

                h_edges_t = h_edges
                v_edges_t = v_edges

                h_edges,v_edges,h_edges_n,v_edges_n = TableFormation(fixed,h_edges + v_edges,table)
                
                
                
                
                found_1,tableD,table,h_edges,v_edges,idx = TableDetection(fixed,h_edges,v_edges,{},index)
                
                
                if not found_1:
                    
                    if tableD != None:
                        tableData.append(tableD)
                        
                    if table != None:
                        tableList.append(table)
                    break
                else:
                    h_edges = h_edges_t
                    v_edges = v_edges_t
                
            h_edges = h_edges+h_edges_n
            v_edges = v_edges+v_edges_n
                    
        index = index + 1        
                    
     return tableData,tableList,indexList2

def play_one(page_idx,images):
    """Run the full pipeline for (page_idx, tolerance-set #cfg_idx).
       Return True if a table was formed, else False."""
    # Apply tolerances for this worker
   
    

    # Load ONLY this page (pdf2image is 1-based)
    
    img = np.array(images[page_idx])
    

    final, angle, conf = full_process(img)
    fixed, angle, conf = deskew(img, angle)
    

    # Detect
   
    
    
    output_image, h_edges, v_edges = EdgeDetection(fixed)
   
    return h_edges,v_edges,fixed


def TableDetection(fixed,h_edges,v_edges,oldList,index):
    # Detects a table
    with open(BASE_DIR / "tolerances" / "tolerances_1.json", "r") as f:
        tolerance_configs = json.load(f)
    
    if len(h_edges) != 0 and len(v_edges) != 0:
        h_edges_o = [edge.copy() for edge in h_edges]
        v_edges_o = [edge.copy() for edge in v_edges]
        it = 0
        if index < len(oldList) and oldList[index]:
            it = oldList[index]
       
        for i, tolerance in enumerate(tolerance_configs[it:]):
            diff = it-i
            cfg = tolerance_configs[i]
            d   = cfg["DEFAULT_TOLERANCE"]
            ex  = cfg["EXTENSION_TOLERANCE"]
            ed  = cfg["EDGES_TOLERANCE"]
            g   = cfg["GAP_TOLERANCE"]
            sx  = cfg["SNAPX_TOLERANCE"]
            sy  = cfg["SNAPY_TOLERANCE"]
            jx  = cfg["JOINX_TOLERANCE"]
            jy  = cfg["JOINY_TOLERANCE"]
            pg  = cfg["PAGE_TOLERANCE"]
                        
            config.DEFAULT_TOLERANCE = d
            config.EXTENSION_TOLERANCE = ex
            config.EDGES_TOLERANCE = ed
            config.GAP_TOLERANCE = g
            config.SNAP_X_TOLERANCE = sx
            config.SNAP_Y_TOLERANCE = sy
            config.JOIN_X_TOLERANCE = jx
            config.JOIN_Y_TOLERANCE = jy
            page.PAGE_TOLERANCE = pg

                        
            h, w = fixed.shape[:2]
            h_edges, v_edges= remove_page_border_edges(h_edges, v_edges, w, h)
                        
            formed,found,h_edges,v_edges,tableD,table,isValid = TableWriter(
                h_edges, v_edges, fixed
                )
            if formed:
                return found,tableD,table,h_edges,v_edges,isValid,i+diff
            
            else:
                h_edges = [edge.copy() for edge in h_edges_o]
                v_edges = [edge.copy() for edge in v_edges_o]
            

    return found,tableD,table,h_edges,v_edges,isValid,maxTolerances

def ArrangeTableFormed(tableData,tableList,fixed,h_edges,v_edges,indexList):
    # Arranges tables after tables are detected
    isFormed = True
    while True:
        tables = CheckRearrangeTables(tableData)
        try:
            sheetData,lindex1,lindex2,isFormed = ArrangeTables(tables)
            isDone = True
            break
        except:
            if len(tableData) > 0:
                tableData.clear()
            if len(tableList) > 0:
                tableList.clear()
            if all(v == maxTolerances for v in indexList.values()):
                break
            
            tableData,tableList,indexList = TableDetectionComp(h_edges,v_edges,tableData,tableList,fixed,indexList)
            
            
    return isDone,sheetData,indexList,lindex1,lindex2,isFormed

def MaxTable(tables):
    # Finds the maximum table out of all tables formed
    max_height = 0
    max_table = tables[0]
    for table in tables:
        height = table[3]["top"] - table[2]["top"]
        if height > max_height:
            max_table = table
            max_height = height
    
    return max_table

def RemoveSmallTables(tables,maxTable):
    # Removes small tables that are relatively smaller compared with the maximun table formed
    heightTo = maxTable[3]["top"] - maxTable[2]["top"]
    widthTo =  maxTable[1]["x0"] - maxTable[0]["x0"]
    tablesToRemove = []
    for table in tables:
        height = table[3]["top"] - table[2]["top"]
        width = table[1]["x0"] - table[0]["x0"]
        if height <= heightTo*0.15 and width <= widthTo*0.15:
            tablesToRemove.append(table)
    if tablesToRemove:
        tables = [e for e in tables if e not in tablesToRemove]

    return tablesToRemove               

def run(pdf_path,on_page_done):
    
    # Returns excel from scanned pdf tables
   
    wb = Workbook()
    images = pdf_to_images(pdf_path, dpi=300)
    
    num_pages = len(images)
    
    page_idx = 0
    
    
    while page_idx < num_pages:
        
        lindex1 = 1
        lindex2 = 1
        rawText = []
        allText = []
        textGroupedXe = {}
        textGroupedXm = {}
        textGroupedXs = {}
        textGrouped = {}
        textGrouped_o = {}
        if page_idx == 0:
            
            ws = wb.active
        else:
            ws = wb.create_sheet()
        try:
            tableData = []
            tableList = []
            sheetData = []
            isFormed = True
            tableList_r = []
            
            isDone = False
            notValid = []
            h_edges,v_edges,fixed = play_one(page_idx,images)
            
            h_edges_o = copy.deepcopy(h_edges)
            v_edges_o = copy.deepcopy(v_edges)
                
            tableData,tableList,indexList = TableDetectionComp(h_edges,v_edges,tableData,tableList,fixed,{})

            
            
            if tableList:
                maxTable = MaxTable(tableList)
                tableList_r = RemoveSmallTables(tableList,maxTable)
            
        
            
            if tableList:
                for tb in tableData[:]:
                    (h_edges_tb_or,v_edges_tb_or,h_edges_or,v_edges_or,tableComp_or,table_or,max_dim_or) = tb
                    if table_or in tableList_r[:]:
                        tableData.remove(tb)
                        tableList.remove(table_or)


            if tableList_r: 
                
                isDone,sheetData,indexList,lindex1,lindex2,isFormed = ArrangeTableFormed(tableData,tableList,fixed,h_edges_o,v_edges_o,{})  
            elif tableList:
                isDone,sheetData,indexList,lindex1,lindex2,isFormed = ArrangeTableFormed(tableData,tableList,fixed,h_edges_o,v_edges_o,indexList) 
            
                            
            
            rawText,textGroupedXm,textGroupedXs,textGorupedXe,textGrouped, image_bgr = TextDetection(fixed)
            
            
            if isDone:
                
                while True:
                    try:
                        
                        wb,sum,count,notValid,textGrouped = WriteText(sheetData,textGrouped,wb,ws)
                        
                        if notValid:
                            
                            for sheet_idx, sheet in enumerate(notValid):
                                index1, index2, tableComp_1, table_o, max_dim = sheet

                                for table in tableData[:]:   # iterate over a COPY
                                    h_edges_tb, v_edges_tb, h_edges, v_edges, tableComp_2, table_o_1, max_dim_o = table

                                    if tableComp_2 == tableComp_1:
                                        tableData.remove(table)
                                    if table_o == table_o_1:
                                        tableList.remove(table_o)
                            
                            isDone2,sheetData,indexList,lindex1,lindex2,isFormed = ArrangeTableFormed(tableData,tableList,fixed,h_edges_o,v_edges_o,{})
                            
                            wb,sum,count,notValid,textGrouped = WriteText(sheetData,textGrouped,wb,ws)
                        
                        
                        break
                    except:
                        if len(tableData) > 0:
                            tableData.clear()
                        if len(tableList) > 0:
                            tableList.clear()
                        for v in indexList.values():
                            print(v)
                        if all(v == maxTolerances for v in indexList.values()):
                                
                                break
                        
                        tableData,tableList,indexList = TableDetectionComp(h_edges_o,v_edges_o,tableData,tableList,fixed,indexList)
                        
                        isDone1,sheetData,indexList,lindex1,lindex2,isFormed = ArrangeTableFormed(tableData,tableList,fixed,h_edges_o,v_edges_o,indexList)
                        
                        
            
                    
        except:
           
               rawText,textGroupedXm,textGroupedXs,textGorupedXe,textGrouped, image_bgr = TextDetection(fixed) 

        finally:
            try:
                if not isFormed:
                      rawText,textGroupedXm,textGroupedXs,textGorupedXe,textGrouped, image_bgr = TextDetection(fixed)    
                allText = NewText(rawText,textGrouped)
                
                tablesText = TableDetectionText(allText)
                
                TableWriterText(tablesText,wb,ws,lindex1,lindex2)
            except:
                pass

        if on_page_done:
            on_page_done(page_idx+1,num_pages)
                    
        page_idx = page_idx + 1

    

    return wb
            

if __name__ == "__main__":
    run(pdf_path,on_page_done)
    
