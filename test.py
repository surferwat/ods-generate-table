# Tests the script on template table and template input forms 
# populated with example values.

import unittest
from unittest import mock
import os
from main import Config, Table, TableEntries





class TestTableEntries(unittest.TestCase):
  @classmethod
  @mock.patch.dict("os.environ", {"PATH_TO_SOURCE_FILES": "test/screened", "PATH_TO_TABLE_TEMPLATE": "test/test_table.ods" })
  def setUpClass(cls):
    cls.config = Config()
    cls.table_entries = TableEntries(cls.config.path_to_source_files, cls.config.subject_month)
    
  def tearDown(self):
    self.table_entries.source_file_names = []
  
  def test_set_subject_month(self):
    month = "08"
    self.table_entries.set_subject_month(month)
    self.assertEqual(self.table_entries.subject_month,"08")

  def test_get_data(self):
    file_name = "test/screened/20220811_test.ods"
    sheet = self.table_entries.get_data(file_name)
    self.assertEqual(sheet['A1'], "スクリーニングフォーム")

  def test_file_filter_condition(self):
    file_name = "test/screened/20220811_test.ods"
    condition = self.table_entries.file_filter_condition(file_name)
    self.assertTrue(condition)
    file_name = "test/screened/20220811_test.odt"
    condition = self.table_entries.file_filter_condition(file_name)
    self.assertFalse(condition)

  def test_check_dir_path(self):
    dir_path = "."
    self.table_entries.check_dir_path(dir_path)
  
  def test_extract_month(self):
    file_path = "test/screened/20220811_test.ods"
    month = self.table_entries.extract_month(file_path)
    self.assertEqual(month, "08")

  def test_set_source_file_names(self):
    dir_path = "test/screened"
    self.table_entries.set_source_file_names(dir_path)
    self.assertEqual(self.table_entries.source_file_names, ["test/screened/20220811_test.ods"])
  
  def test_set_total_area(self):
    building_area = "-"
    room_area = "37"
    total_area = self.table_entries.set_total_area(building_area, room_area)
    self.assertEqual(total_area,"37")
    
  def test_generate(self):
    dir_path = "test/screened"
    self.table_entries.set_source_file_names(dir_path)
    entries = self.table_entries.generate()
    self.assertEqual(len(entries), 1)





class TestTable(unittest.TestCase):
  @classmethod
  @mock.patch.dict("os.environ", {"PATH_TO_SOURCE_FILES": "test/screened", "PATH_TO_TABLE_TEMPLATE": "test/test_table.ods" })
  def setUpClass(cls):
    cls.config = Config()
    cls.table = Table(cls.config.path_to_table_template)
    cls.table.table_dest_path = "test"
  
  def test_set_table_entries(self):
    entries = [{
      'admin_area': "admin_area", 
      'district': "district", 
      'property_type': "property_type", 
      'year_built': "year_built",
      'total_area': "total_area", 
      'price': "price", 
      'price_per_area': "price_per_area",
      'file_path': "file_path"
    }]
    self.table.set_table_entries(entries)
    self.assertEqual(self.table.table_entries, entries)
  
  def test_check_file_path(self):
    file_path = "test/screened/20220811_test.ods"
    self.table.check_file_path(file_path)
  
  def test_get_data(self):
    file_name = "test/screened/20220811_test.ods"
    sheet = self.table.get_data(file_name)
    self.assertEqual(sheet['A1'], "スクリーニングフォーム")

  def test_generate(self):
    entries = [{
      'admin_area': "admin_area", 
      'district': "district", 
      'property_type': "property_type", 
      'year_built': "year_built",
      'total_area': "total_area", 
      'price': "price", 
      'price_per_area': "price_per_area",
      'file_path': "file_path"
    }]
    self.table.set_table_entries(entries)
    self.table.generate()
    self.assertEqual(self.table.table_sheet[4,0],1)
    self.assertEqual(self.table.table_sheet[4,1],"admin_area")
    self.assertEqual(self.table.table_sheet[4,2],"district")
    self.assertEqual(self.table.table_sheet[4,3],"property_type")
    self.assertEqual(self.table.table_sheet[4,4],"year_built")
    self.assertEqual(self.table.table_sheet[4,5],"total_area")
    self.assertEqual(self.table.table_sheet[4,6],"price")
    self.assertEqual(self.table.table_sheet[4,7],"price_per_area")
  
  def test_save(self):
    entries = [{
      'admin_area': "admin_area", 
      'district': "district", 
      'property_type': "property_type", 
      'year_built': "year_built",
      'total_area': "total_area", 
      'price': "price", 
      'price_per_area': "price_per_area",
      'file_path': "file_path", 
    }]
    self.table.set_table_entries(entries)
    self.table.generate()
    self.table.save()
    self.assertTrue(os.path.isfile(self.table.table_dest_path + "/" + self.table.get_date() + "_" + "activity_table.ods"))





if __name__ == '__main__':
  unittest.main()