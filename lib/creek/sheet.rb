require 'zip/filesystem'
require 'nokogiri'

module Creek
  class Creek::Sheet
    include Creek::Utils

    attr_reader :book,
                :name,
                :sheetid,
                :state,
                :visible,
                :rid,
                :index,
                :headers

    # An XLS file has only 256 columns, however, an XLSX or XLSM file can contain up to 16384 columns.
    # This function creates a hash with all valid XLSX column names and associated indices.
    # Note: load and memoize on demand
    def self.column_indexes
      @column_indexes ||= ('A'..'XFD').each_with_index.to_h.freeze
    end

    def initialize(book, name, sheetid, state, visible, rid, sheetfile)
      @book = book
      @name = name
      @sheetid = sheetid
      @visible = visible
      @rid = rid
      @state = state
      @sheetfile = sheetfile
      @images_present = false
    end

    def column_indexes
      self.class.column_indexes
    end

    ##
    # Preloads images info (coordinates and paths) from related drawing.xml and drawing rels.
    # Must be called before #rows method if you want to have images included.
    # Returns self so you can chain the calls (sheet.with_images.rows).
    def with_images
      @drawingfile = extract_drawing_filepath
      if @drawingfile
        @drawing = Creek::Drawing.new(@book, @drawingfile.sub('..', 'xl'))
        @images_present = @drawing.has_images?
      end
      self
    end

    ##
    # Extracts images for a cell to a temporary folder.
    # Returns array of Pathnames for the cell.
    # Returns nil if images asre not found for the cell or images were not preloaded with #with_images.
    def images_at(cell)
      @drawing.images_at(cell) if @images_present
    end

    ##
    # Provides an Enumerator that returns a hash representing each row.
    # The key of the hash is the Cell id and the value is the value of the cell.
    def rows(headers: false, header_row_number: 1, metadata: false)
      extract_headers(header_row_number) if headers

      rows_generator(include_headers: headers, include_meta_data: metadata)
    end

    ##
    # Provides an Enumerator that returns a hash representing each row.
    # The hash contains meta data of the row and a 'cells' embended hash which contains the cell contents.
    def rows_with_meta_data(headers: false, header_row_number: 1)
      rows(headers: headers, header_row_number: header_row_number, metadata: true)
    end

    # Parses the file until the header row is reached.
    # Returns the headers as an array.
    def extract_headers(row_number = 1)
      return @headers if defined?(@headers)

      # Extracted row numbers are String, convert it here to facilite comparison
      @header_row_number = row_number.to_s

      rows_with_meta_data.each do |row|
        return (@headers = row['cells'].any? && row['cells']) if @header_row_number == row['r']
      end
    end

    private

    CELL = 'c'.freeze
    ROW = 'row'.freeze
    VALUE = 'v'.freeze
    CELL_TYPE = 't'.freeze
    STYLE_INDEX = 's'.freeze
    CELL_REF = 'r'.freeze
    ROW_NUMBER = 'r'.freeze
    TEXT = 't'.freeze

    ##
    # Returns an array or hash (with headers as key) per row that includes the cell ids and values.
    # Empty cells will be also included with a nil value.
    def rows_generator(include_meta_data: false, include_headers: false)
      path =
        if @sheetfile.start_with?("/xl/") || @sheetfile.start_with?("xl/")
          @sheetfile
        else
          "xl/#{@sheetfile}"
        end

      if @book.files.file.exist?(path)
        # SAX parsing, Each element in the stream comes through as two events:
        # one to open the element and one to close it.
        opener = Nokogiri::XML::Reader::TYPE_ELEMENT
        closer = Nokogiri::XML::Reader::TYPE_END_ELEMENT

        Enumerator.new do |y|
          row, cells, cell = nil, [], nil
          row_number = nil
          cell_type  = nil
          cell_style_idx = nil

          @book.files.file.open(path) do |xml|
            Nokogiri::XML::Reader.from_io(xml).each do |node|
              node_name = node.name
              next if node.node_type != opener && node_name != ROW

              if node_name == ROW
                case node.node_type
                when opener
                  row = node.attributes
                  row_number = row[ROW_NUMBER]

                  if (spans = row['spans'])
                    spans = spans.split(":").last.to_i - 1
                  else
                    spans = 0
                  end

                  cells = Array.new(spans)

                  if node.self_closing?
                    y << to_formatted_row(row, cells, include_meta_data, include_headers)
                  end
                when closer
                  y << to_formatted_row(row, cells, include_meta_data, include_headers)
                end
              elsif node_name == CELL
                attributes = node.attributes
                cell_type      = attributes[CELL_TYPE]
                cell_style_idx = attributes[STYLE_INDEX]
                cell           = attributes[CELL_REF]
              elsif node_name == VALUE
                if cell
                  cells[column_indexes[cell.sub(row_number, '')]] = convert(node.inner_xml, cell_type, cell_style_idx)
                end
              elsif node_name == TEXT
                if cell
                  cells[column_indexes[cell.sub(row_number, '')]] = convert(node.inner_xml, cell_type, cell_style_idx)
                end
              end
            end
          end
        end
      end
    end

    def to_formatted_row(row, cells, include_meta_data, include_headers)
      if include_headers
        row['header_row'] = row[ROW_NUMBER] == @header_row_number
        cells = cells_with_headers(cells) if @headers
      end

      if include_meta_data
        row['cells'] = cells
        row
      else
        cells
      end
    end

    def cells_with_headers(cells)
      cells.empty? ? {} : @headers.zip(cells).to_h
    end

    def convert(value, type, style_idx)
      style = @book.style_types[style_idx.to_i]
      Creek::Styles::Converter.call(value, type, style, converter_options)
    end

    def converter_options
      @converter_options ||= { shared_strings: @book.shared_strings.dictionary }
    end

    ##
    # Find drawing filepath for the current sheet.
    # Sheet xml contains drawing relationship ID.
    # Sheet relationships xml contains drawing file's location.
    def extract_drawing_filepath
      # Read drawing relationship ID from the sheet.
      sheet_filepath = "xl/#{@sheetfile}"
      drawing = parse_xml(sheet_filepath).css('drawing').first
      return if drawing.nil?

      drawing_rid = drawing.attributes['id'].value

      # Read sheet rels to find drawing file's location.
      sheet_rels_filepath = expand_to_rels_path(sheet_filepath)
      parse_xml(sheet_rels_filepath).css("Relationship[@Id='#{drawing_rid}']").first.attributes['Target'].value
    end
  end
end
