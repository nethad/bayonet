module Bayonet
  class Cell

    ROW_MATCHER    = /\A[A-Z]+\z/ # is only A-Z characters, upper case.

    def initialize(row, column)
      @row    = row
      @column = column
    end

    def label
      cell_label = "#{row}#{column}"
      if valid?
        cell_label
      else
        raise "Invalid cell: #{cell_label}."
      end
    end

    def valid?
      row_valid? && column_valid?
    end

    private

    attr_reader :row, :column

    def row_valid?
      (row =~ ROW_MATCHER) != nil
    end

    def column_valid?
      column.is_a?(Integer) && column >= 1
    end

  end
end
