module Bayonet
  class NodeCreation

    attr_reader :cell_label, :sheet

    def initialize(cell_label, sheet)
      @cell_label = cell_label
      @sheet      = sheet
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
      @cell_node ||= xml.at_css("c[r=\"#{cell_label}\"]")
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
      @row_number ||= cell_label.gsub(/[^\d]/, '')
    end

    def sheet_data_node
      @sheet_data_node ||= xml.at_css('sheetData')
    end

    def create_cell_node
      create_node('c', row_node, cell_label)
    end

    def create_row_node
      create_node('row', sheet_data_node, row_number)
    end

    def create_node(tag, parent_node, value)
      Nokogiri::XML::Node.new(tag, parent_node).tap do |node|
        node['r'] = value
        parent_node.add_child(node)
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
