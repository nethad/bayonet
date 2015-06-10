module Bayonet
  class Sheet

    def initialize(name, workbook)
      @name = name
      @workbook = workbook
    end

    def path
      "xl/worksheets/sheet#{id}.xml"
    end

    def xml
      @xml ||= read_and_parse_xml
    end

    def write_string(cell_row, cell_column, value)
      set_cell(cell_row, cell_column, value.to_s, :str)
    end

    def write_number(cell_row, cell_column, value)
      set_cell(cell_row, cell_column, value, :n)
    end

    def set_cell(cell_row, cell_column, value, type = nil)
      cell_node = get_or_create_cell_node(Cell.new(cell_row, cell_column))
      cell_node['t'] = type unless type.nil?
      set_value(cell_node, value)
    end

    def set_typed_cell(cell_row, cell_column, value)
      if (value.is_a?(Numeric))
        write_number(cell_row, cell_column, value)
      else
        write_string(cell_row, cell_column, value)
      end
    end

    private

    attr_reader :name, :workbook

    def id
      @id ||= workbook.xml.at_css("sheets sheet[name=\"#{name}\"]")["r:id"][3..-1]
    end

    def read_and_parse_xml
      entry = workbook.zip_file.find_entry(path)
      Nokogiri::XML.parse(entry.get_input_stream)
    end

    def get_or_create_cell_node(cell)
      Bayonet::NodeCreation.new(cell.label, self).get_or_create_cell_node
    end

    def set_value(cell_node, value)
      if cell_node.children.nil? || cell_node.children.empty?
        value_node = Nokogiri::XML::Node.new('v', cell_node)
        value_node.content = value
        cell_node.add_child value_node
      else
        cell_node.at_css("v").content = value
      end
    end
  end
end
