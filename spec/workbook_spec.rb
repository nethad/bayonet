require 'spec_helper'
require 'roo'

describe Bayonet::Workbook do

  let(:workbook) { Bayonet::Workbook.new(File.new('spec/fixtures/fixture1.xlsx')) }

  it 'writes a string to an existing cell' do
    workbook.on_sheet('The Sheep') do |sheet|
      sheet.set_typed_cell('A', 1, 'Yay')
    end

    Tempfile.open('output.xlsx') do |file|
      workbook.write_and_close(file.path)

      roo_xlsx = Roo::Excelx.new(file.path, :packed => false, :file_warning => :ignore)

      sheet = roo_xlsx.sheet('The Sheep')
      expect(sheet.cell('A', 1)).to eq('Yay')
    end
  end

  it 'write a string to an existing row, but non-existing cell' do
    workbook.on_sheet('The Sheep') do |sheet|
      sheet.set_typed_cell('A', 1, 'Yay')
      sheet.set_typed_cell('B', 1, 'Yay2')
    end

    Tempfile.open('output.xlsx') do |file|
      workbook.write_and_close(file.path)

      roo_xlsx = Roo::Excelx.new(file.path, :packed => false, :file_warning => :ignore)

      sheet = roo_xlsx.sheet('The Sheep')
      expect(sheet.cell('A', 1)).to eq('Yay')
      expect(sheet.cell('B', 1)).to eq('Yay2')
    end
  end

  it 'write a string to an non-existing row and cell' do
    workbook.on_sheet('The Sheep') do |sheet|
      sheet.set_typed_cell('A', 1, 'Yay')
      sheet.set_typed_cell('A', 2, 'Yay2')
    end

    Tempfile.open('output.xlsx') do |file|
      workbook.write_and_close(file.path)

      roo_xlsx = Roo::Excelx.new(file.path, :packed => false, :file_warning => :ignore)

      sheet = roo_xlsx.sheet('The Sheep')
      expect(sheet.cell('A', 1)).to eq('Yay')
      expect(sheet.cell('A', 2)).to eq('Yay2')
    end
  end

  context 'multiple writes' do

    it 'writes a string to an existing cell' do
      workbook.on_sheet('The Sheep') do |sheet|
        sheet.set_typed_cell('A', 1, 'Yay')
      end

      workbook.on_sheet('The Sheep') do |sheet|
        sheet.set_typed_cell('A', 1, 'Second Value')
      end

      Tempfile.open('output.xlsx') do |file|
        workbook.write_and_close(file.path)

        roo_xlsx = Roo::Excelx.new(file.path, :packed => false, :file_warning => :ignore)

        sheet = roo_xlsx.sheet('The Sheep')
        expect(sheet.cell('A', 1)).to eq('Second Value')
      end
    end

    it 'write a string to an existing row, but non-existing cell' do
      workbook.on_sheet('The Sheep') do |sheet|
        sheet.set_typed_cell('A', 1, 'Yay')
        sheet.set_typed_cell('B', 1, 'Yay2')
      end

      workbook.on_sheet('The Sheep') do |sheet|
        sheet.set_typed_cell('A', 1, 'Second Value A1')
        sheet.set_typed_cell('B', 1, 'Second Value B2')
      end

      Tempfile.open('output.xlsx') do |file|
        workbook.write_and_close(file.path)

        roo_xlsx = Roo::Excelx.new(file.path, :packed => false, :file_warning => :ignore)

        sheet = roo_xlsx.sheet('The Sheep')
        expect(sheet.cell('A', 1)).to eq('Second Value A1')
        expect(sheet.cell('B', 1)).to eq('Second Value B2')
      end
    end

    it 'write a string to an non-existing row and cell' do
      workbook.on_sheet('The Sheep') do |sheet|
        sheet.set_typed_cell('A', 1, 'Yay')
        sheet.set_typed_cell('A', 2, 'Yay2')
      end

      workbook.on_sheet('The Sheep') do |sheet|
        sheet.set_typed_cell('A', 1, 'Second Value A1')
        sheet.set_typed_cell('A', 2, 'Second Value A2')
      end

      Tempfile.open('output.xlsx') do |file|
        workbook.write_and_close(file.path)

        roo_xlsx = Roo::Excelx.new(file.path, :packed => false, :file_warning => :ignore)

        sheet = roo_xlsx.sheet('The Sheep')
        expect(sheet.cell('A', 1)).to eq('Second Value A1')
        expect(sheet.cell('A', 2)).to eq('Second Value A2')
      end
    end

  end

end
