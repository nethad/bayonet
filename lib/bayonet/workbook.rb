require 'zip'
require 'nokogiri'

module Bayonet
  class Workbook

    WORKBOOK_PATH = 'xl/workbook.xml'

    def initialize(file)
      @file = file
    end

    def zip_file
      @zip_file ||= Zip::File.open(file)
    end

    def xml
      @xml ||= read_and_parse_xml
    end

    def on_sheet(sheet_name, &block)
      sheet = find_sheet_by_name_and_cache(sheet_name)
      block.call(sheet)
    end

    def write(file)
      Zip::File.open(file, Zip::File::CREATE) do |destination|
        zip_file.each do |entry|
          destination.get_output_stream(entry.name) do |file_stream|
            write_modifications(entry, file_stream)
          end
        end
      end
    end

    def close
      zip_file.close
    end

    def write_and_close(file)
      write(file)
      close
    end

    private

    attr_reader :file

    def find_sheet_by_name_and_cache(sheet_name)
      return modified_sheets[sheet_name] if modified_sheets.key?(sheet_name)

      sheet = Bayonet::Sheet.new(sheet_name, self)
      modified_sheets[sheet_name] = sheet
      sheet
    end

    def find_sheet_by_path(path)
      modified_sheets.values.detect { |sheet| sheet.path == path }
    end

    def write_modifications(entry, file_stream)
      sheet = find_sheet_by_path(entry.name)
      if sheet
        file_stream.write(sheet.xml.to_s)
      else
        file_stream.write(zip_file.read(entry.name))
      end
    end

    def modified_sheets
      @modified_sheets ||= {}
    end

    def read_and_parse_xml
      entry = zip_file.find_entry(WORKBOOK_PATH)
      Nokogiri::XML.parse(entry.get_input_stream)
    end
  end
end
