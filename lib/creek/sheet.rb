require 'zip/filesystem'
require 'nokogiri'

module Creek
  class Creek::Sheet

    attr_reader :book,
                :name,
                :sheetid,
                :state,
                :visible,
                :rid,
                :index


    def initialize book, name, sheetid, state, visible, rid, sheetfile
      @book = book
      @name = name
      @sheetid = sheetid
      @visible = visible
      @rid = rid
      @state = state
      @sheetfile = sheetfile

      # An XLS file has only 256 columns, however, an XLSX or XLSM file can contain up to 16384 columns.
      # This function creates a hash with all valid XLSX column names and associated indices.
      @excel_col_names = Hash.new
      (0...16384).each do |i|
        @excel_col_names[col_name(i)] = i
      end
    end

    ##
    # Provides an Enumerator that returns a hash representing each row.
    # The key of the hash is the Cell id and the value is the value of the cell.
    def rows
      rows_generator
    end

    ##
    # Provides an Enumerator that returns a hash representing each row.
    # The hash contains meta data of the row and a 'cells' embended hash which contains the cell contents.
    def rows_with_meta_data
      rows_generator true
    end

    private
    ##
    # Returns valid Excel column name for a given column index.
    # For example, returns "A" for 0, "B" for 1 and "AQ" for 42.
    def col_name i
      i += 1
      result = ""
      mapping = ('A'..'Z').to_a
      while i > 0
        mod = (i - 1) % 26
        result += mapping[mod]
        i = (i - mod) / 26
      end

      result.reverse
    end

    CELL = 'c'.freeze
    ROW = 'row'.freeze
    VALUE = 'v'.freeze
    CELL_TYPE = 't'.freeze
    STYLE_INDEX = 's'.freeze
    CELL_REF = 'r'.freeze

    ##
    # Returns a hash per row that includes the cell ids and values.
    # Empty cells will be also included in the hash with a nil value.
    def rows_generator include_meta_data=false
      path = "xl/#{@sheetfile}"
      if @book.files.file.exist?(path)
        # SAX parsing, Each element in the stream comes through as two events:
        # one to open the element and one to close it.
        opener = Nokogiri::XML::Reader::TYPE_ELEMENT
        closer = Nokogiri::XML::Reader::TYPE_END_ELEMENT
        Enumerator.new do |y|
          row, cells, cell = nil, {}, nil
          cell_type  = nil
          cell_style_idx = nil
          @book.files.file.open(path) do |xml|
            Nokogiri::XML::Reader.from_io(xml).each do |node|
              node_name = node.name
              next unless node_name == CELL || node_name == ROW || node_name == VALUE
              if node_name == ROW
                case node.node_type
                when opener then
                  row = node.attributes
                  row['cells'] = Hash.new
                  cells = Hash.new
                  y << (include_meta_data ? row : cells) if node.self_closing?
                when closer
                  processed_cells = fill_in_empty_cells(cells, row['r'], cell)
                  row['cells'] = processed_cells
                  y << (include_meta_data ? row : processed_cells)
                end
              elsif (node_name == CELL) && node.node_type == opener
                attributes = node.attributes
                cell_type      = attributes[CELL_TYPE]
                cell_style_idx = attributes[STYLE_INDEX]
                cell           = attributes[CELL_REF]

              elsif node_name == VALUE && node.node_type == opener
                if !cell.nil?
                  cells[cell] = convert(node.inner_xml, cell_type, cell_style_idx)
                end
              end
            end
          end
        end
      end
    end

    def convert(value, type, style_idx)
      style = @book.style_types[style_idx.to_i]
      Creek::Styles::Converter.call(value, type, style, converter_options)
    end

    def converter_options
      @converter_options ||= {shared_strings: @book.shared_strings.dictionary}
    end

    ##
    # The unzipped XML file does not contain any node for empty cells.
    # Empty cells are being padded in using this function
    def fill_in_empty_cells cells, row_number, last_col
      new_cells = Hash.new
      unless cells.empty?
        last_col = last_col.gsub(row_number, '')
        last_col_index = @excel_col_names[last_col]
        [*(0..last_col_index)].each do |i|
          col = col_name i
          id = "#{col}#{row_number}"
          unless cells.has_key? id
              new_cells[id] = nil
          else
            new_cells[id] = cells[id]
          end
        end
      end
      new_cells
    end
  end
end
