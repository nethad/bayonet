require 'spec_helper'
require 'roo'

describe 'cell types' do

  let(:workbook) { Bayonet::Workbook.new(File.new('spec/fixtures/fixture1.xlsx')) }

  it 'writes a string to a cell' do
    workbook.on_sheet('The Sheep') do |sheet|
      sheet.write_string('A1', 'Yay')
    end

    Tempfile.open('output.xlsx') do |file|
      workbook.write(file.path)
      workbook.close

      roo_xlsx = Roo::Excelx.new(file.path, :packed => false, :file_warning => :ignore)

      sheet = roo_xlsx.sheet('The Sheep')
      expect(sheet.cell('A', 1)).to eq('Yay')
    end
  end

  it 'writes a number to a cell' do
    workbook.on_sheet('The Sheep') do |sheet|
      sheet.write_number('A2', 42)
    end

    Tempfile.open('output.xlsx') do |file|
      workbook.write(file.path)
      workbook.close

      roo_xlsx = Roo::Excelx.new(file.path, :packed => false, :file_warning => :ignore)

      sheet = roo_xlsx.sheet('The Sheep')
      expect(sheet.cell('A', 2)).to eq(42)
    end
  end

end