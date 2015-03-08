module Bayonet
  class NodeCreation

    attr_reader :cell, :sheet

    def initialize(cell, sheet)
      @cell  = cell
      @sheet = sheet
    end

    def get_or_create_cell_node
      unless has_cell_node?
        create_nodes
      end

      cell_node
    end

    private

    def xml
      sheet.xml
    end

    def cell_node
      @cell_node ||= xml.at_css("c[r=\"#{cell}\"]")
    end

    def has_cell_node?
      !cell_node.nil?
    end

    def row_node
      @row_node ||= xml.at_css("row[r=\"#{row_number}\"]")
    end

    def has_row_node?
      !row_node.nil?
    end

    def row_number
      @row_number ||= cell.gsub(/[^\d]/, '')
    end

    def sheet_data_node
      @sheet_data_node ||= xml.at_css('sheetData')
    end

    def create_cell_node
      Nokogiri::XML::Node.new('c', row_node).tap do |cell_node|
        cell_node['r'] = cell
        row_node.add_child(cell_node)
      end
    end

    def create_row_node
      Nokogiri::XML::Node.new('row', sheet_data_node).tap do |row_node|
        row_node['r'] = row_number
        sheet_data_node.add_child(row_node)
      end
    end

    def create_row_node_with_cell_node
      row_node  = create_row_node
      cell_node = create_cell_node
      row_node.add_child(cell_node)
    end

    def create_nodes
      if has_row_node?
        create_cell_node
      else
        create_row_node_with_cell_node
      end
    end

  end
end
