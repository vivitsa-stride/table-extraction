# convert json to input object
from extractTables import Extract
# from extract_tables_from_bbox import extract_tables as bbox_extract

class TableProcessing():
    def __init__(self, input_json):
        self.input_json = input_json
        self._create_table_input()
        print(self.tables)
            
        self.resultFiles = self._get_table_output()

    def _create_table_input(self):
        self.filePath = self.input_json["filePath"]
        self.tables = self.input_json["values"]        

    def _get_table_output(self):
        resFilepaths = []
        e = Extract(self.filePath,self.tables)
        resFilepath = e.resFilepath
        resFilepaths.append(resFilepath)
        return resFilepaths



    # pdf--pdf
    # (img)png,jpg,jpeg--ocrization--pdf
    # if pdf - multiprocessing pagewise extraction-special case multipagers
    # input fields given by user will decide what type of extraction
    # - user could give page numbr, bbox or not
    # fields to be entered:
    # only one file-filename
    # None/[page number, None/bbox]
    # if filename, None: generic --- extract all from all pages
    # if filename, page_number, None: semi generic --- extract all from given page only
    # if filename, page_number, bbox: targeted --- extract all from given page number and bbox

    

    def _ocrize_img_to_pdf(self):
        pass
      
    

    
        
        


        


input_json = {
  "filePath": "/home/vaibhav/Downloads/Moodys Sfg2/batch1/9193_20190620.pdf",
  "values": {
    "value1": {
      "pageNum": 1,
      "bbox_value": [],
    },
    "value2": {
      "pageNum": None,
      "bbox_value": None,
    },
    "value3": {
      "pageNum":None,
      "bbox_value": [],
    },
  },
}

i = TableProcessing(input_json)
print(i.resultFiles)
