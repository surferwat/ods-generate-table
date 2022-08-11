# ods-generate-table is a CLI script that generates an ODS format 
# (ods extension) table aggregating data from input forms.

import os
import sys
import pyexcel as pe
from datetime import date






class Config:
  path_to_source_files = ""
  path_to_table_template = ""
  subject_month = ""

  def __init__(self):
    try:
      self.path_to_source_files = "test/screened" if os.environ["PATH_TO_SOURCE_FILES"] == "" else os.environ["PATH_TO_SOURCE_FILES"]
    except:
      print("env variable does not exist: PATH_TO_SOURCE_FILES")
    
    try:
      self.path_to_table_template = "test/test_table.ods" if os.environ["PATH_TO_TABLE_TEMPLATE"] == "" else os.environ["PATH_TO_TABLE_TEMPLATE"]
    except:
      print("env variable does not exist: PATH_TO_TABLE_TEMPLATE")

    prompt = "What is the month for which you are generating the table? Enter two digs (e.g., 01 for Jan) "
    self.subject_month = input(prompt) # add validation






class TableEntries:
  # Constants used to identify relevant source files
  FORM_IDENTIFIER = "スクリーニングフォーム"
  FORM_IDENTIFIER_CELL = "A1"
  
  # Constants for locations of values that make up a table entry
  ADMIN_AREA_CELL = 'C25'
  DISTRICT_CELL = 'C13'
  PROPERTY_TYPE_CELL = 'C31'
  YEAR_BUILT_CELL = 'C32'
  BUILDING_AREA_CELL = 'C35'
  ROOM_AREA_CELL = 'C41'
  PRICE_CELL = 'C7'
  PRICE_PER_AREA_CELL = 'C8'

  source_file_names = []
  subject_month = ""
  table_entries = []

  def __init__(self, dir_path, subject_month):
    """Initializes Table Entries.
    Args:
      dir_path: The path to the directory location of source files.
      subject_month: The month that is relevant to the table.
    Returns:
      Nothing.
    """
    self.set_source_file_names(dir_path)
    self.set_subject_month(subject_month)
  
  def set_subject_month(self, subject_month):
    """Sets subject_month with a value representing the month relevant 
    for the table.
    Args:
      subject_month: The month relevant for the table.
    Returns:
      Nothing.
    """
    self.subject_month = subject_month

  def check_dir_path(self, dir_path):
    """Checks whether directory exists.
    Args:
      dir_path: The path to the directory to check.
    Returns:
      Nothing.
    """
    if not os.path.isdir(dir_path):
      print("Dir path does not exist: " + dir_path + ". Program exited.")
      sys.exit(0)
  
  def extract_month(self, file_path):
    """Extracts the two digits that represent the month from the file name.
    Args:
      file_path: The path to the file to be checked.
    Returns:
      Two digits that represent the month in the year.
    """
    pos = file_path.rfind("/")
    month = file_path[pos+5:pos+7]
    if not month.isnumeric():
      print("Invalid file format: ", file_path)
    return month

  def file_filter_condition(self, file_path):
    """Determines whether subject file is a relevant source file to 
    be transformed into a table entry.
    Args:
      file_path: The path to the file to be checked. e.g., _test/screened/test.ods_
    Returns:
      True if relevant source file, False otherwise.
    """
    len_from_end = 3
    month = self.extract_month(file_path)
    isOds = file_path[-len_from_end:] == "ods"
    isRelevantMonth = month == self.subject_month
    if isOds and isRelevantMonth:
      sheet = self.get_data(file_path)
      return sheet[self.FORM_IDENTIFIER_CELL] == self.FORM_IDENTIFIER
    else:
      print("Invalid file format.")
      return False

  def set_source_file_names(self, dir_path):
    """Sets source_file_names with a list of the file names for the 
    relevant source files.
    Args:
      dir_path: The path to the directory where the source files are 
      saved.
    Returns:
      Nothing.
    """
    self.check_dir_path(dir_path)
    for root, dir_names, file_names in os.walk(dir_path):
      for name in file_names:
        file_path = os.path.join(root,name) 
        if self.file_filter_condition(file_path):
          self.source_file_names.append(file_path)
  
  def get_data(self, file_path):
    """Extracts data from sheet.
    Args:
      file_path: The path to the file to extract data from.
    Returns:
      A sheet object.
    """
    return pe.get_sheet(file_name=file_path)
  
  def set_total_area(self, building_area, room_area):
    """Sets a value for the total area depending on the input values for 
    building_area and room_area.
    Args:
      building_area: The total area of the building.
      room_area: The total area of the room.
    Returns:
      A value for the applicable total area.
    """
    if building_area != "-":
      total_area = building_area
    elif room_area != "-":
      total_area = room_area
    else:
      total_area = "-"
    return total_area
    
  def generate(self):
    """Generates a list of dictionaries representing tablerow entries.
    Returns:
      The list of dictionaries.
    """
    for file_path in self.source_file_names:
      sheet = self.get_data(file_path)
  
      admin_area = sheet[self.ADMIN_AREA_CELL]
      district = sheet[self.DISTRICT_CELL]
      property_type = sheet[self.PROPERTY_TYPE_CELL]
      year_built = sheet[self.YEAR_BUILT_CELL]
      total_area = self.set_total_area(sheet[self.BUILDING_AREA_CELL], sheet[self.ROOM_AREA_CELL])
      price = sheet[self.PRICE_CELL]
      price_per_area = sheet[self.PRICE_PER_AREA_CELL]
      self.table_entries.append({
        'admin_area': admin_area, 
        'district': district, 
        'property_type': property_type, 
        'year_built': year_built,
        'total_area': total_area, 
        'price': price, 
        'price_per_area': price_per_area,
        'file_path': file_path
      })
    return self.table_entries





class Table:
  # Constants for the sheet locations of the table header
  # and footer 
  HEADER_COL = 1
  FOOTER_NOTE_COL = 1

  FIRST_ROW = 5
  
  # Constants for the column locations of each column of the table
  NUMBER_COL = 1
  PROPERTY_TYPE_COL = 4
  ADMIN_AREA_COL = 2
  DISTRICT_COL = 3
  YEAR_BUILT_COL = 5
  TOTAL_AREA_COL = 6
  PRICE_COL = 7
  PRICE_PER_AREA_COL = 8
  FILE_PATH_COL = 15
  
  table_template_path = ""
  table_dest_path = "/"
  table_entries = []
  table_sheet = ""

  def __init__(self, file_path):
    """Initiliazes Table
    Args:
      file_path: The path to the table template file.
    Returns:
      Nothing.
    """
    self.check_file_path(file_path)
    self.table_template_path = file_path
  
  def check_file_path(self, file_path):
    """Checks whether file exists.
    Args:
      file_path: The path to the file to check.
    Returns:
      Nothing.
    """
    if not os.path.isfile(file_path):
      print("File path does not exist: " + file_path + ". Program exited.")
      sys.exit(0)
    
  def set_table_dest_path(self):
    """Sets table_dest_path with a string for the path to where
    the generated table should be saved.
    Returns:
      Nothing.
    """
    prompt = "What is the filepath where table should be saved? "
    path = input(prompt)
    while not os.path.isdir(path):
      print("Invalid file path. Input valid file path.")
      path = input(prompt)
    self.table_dest_path = path

  def set_table_entries(self, entries):
    """Sets table_entries with a list of dictionaries representing 
    the row entires of the table.
    Args:
      entries:  The list of dictionaries.
    Returns:
      Nothing.
    """
    self.table_entries = entries
  
  def get_data(self, file_name):
    """Extracts data from sheet.
    Args:
      file_path: The path to the file to extract data from.
    Returns:
      A sheet object.
    """
    return pe.get_sheet(file_name=file_name)

  def generate(self):
    """Populates the table with values for the table entries, header,
    and footer.
    Returns:
      Nothing.
    """
    self.table_sheet = self.get_data(self.table_template_path)
    row = self.FIRST_ROW
    for entry in self.table_entries:
      number = row - self.FIRST_ROW + 1
      self.table_sheet[row - 1, self.NUMBER_COL - 1] = number
      self.table_sheet[row - 1, self.PROPERTY_TYPE_COL - 1] = entry['property_type']
      self.table_sheet[row - 1, self.ADMIN_AREA_COL - 1] = entry['admin_area']
      self.table_sheet[row - 1, self.DISTRICT_COL - 1] = entry['district']
      self.table_sheet[row - 1, self.YEAR_BUILT_COL - 1] = entry['year_built']
      self.table_sheet[row - 1, self.TOTAL_AREA_COL - 1] = entry['total_area']
      self.table_sheet[row - 1, self.PRICE_COL - 1] = entry['price']
      self.table_sheet[row - 1, self.PRICE_PER_AREA_COL - 1] = entry['price_per_area']
      self.table_sheet[row - 1, self.FILE_PATH_COL - 1] = entry['file_path']
      row = row + 1
    
    self.table_sheet[self.FIRST_ROW - 3, self.HEADER_COL - 1] = str(date.today().year) + "年" + str(date.today().month) + "月"
    self.table_sheet[row - 1, self.FOOTER_NOTE_COL - 1] = "注: S=スクリーニング、FR=ファーストレビュー、 PC=予備的結論、C=結論; 「-」＝実施されていない、「o」＝肯定、「ｘ」＝否定。"

  def get_date(self):
    """Gets the current date in YYYYMMDD formatting.
    Returns: 
      Formatted date.
    """
    return date.today().strftime('%Y%m%d')

  def save(self):
    """Saves the table.
    Returns:
      Nothing.
    """
    self.table_sheet.save_as(self.table_dest_path + "/" + self.get_date() + "_" + "activity_table.ods")





if __name__ == "__main__":
  # Set config
  config = Config()

  # Generate table entires
  table_entries = TableEntries(config.path_to_source_files, config.subject_month)
  entries = table_entries.generate()

  # Generate table
  table = Table(config.path_to_table_template)
  table.set_table_dest_path()
  table.set_table_entries(entries)
  table.generate()
  table.save()

  print("done!")