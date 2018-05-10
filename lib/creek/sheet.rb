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
                :index


    def initialize book, name, sheetid, state, visible, rid, sheetfile
      @book = book
      @name = name
      @sheetid = sheetid
      @visible = visible
      @rid = rid
      @state = state
      @sheetfile = sheetfile
      @images_present = false

      # An XLS file has only 256 columns, however, an XLSX or XLSM file can contain up to 16384 columns.
      # This function creates a hash with all valid XLSX column names and associated indices.
      @excel_col_names = Hash.new
      ('A'..'XFD').each_with_index do |col_name, index|
        @excel_col_names[col_name] = index
      end
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

    CELL = 'c'.freeze
    ROW = 'row'.freeze
    VALUE = 'v'.freeze
    CELL_TYPE = 't'.freeze
    STYLE_INDEX = 's'.freeze
    CELL_REF = 'r'.freeze
    ROW_NUMBER = 'r'.freeze
    TEXT = 't'.freeze

    ##
    # Returns a hash per row that includes the cell ids and values.
    # Empty cells will be also included in the hash with a nil value.
    def rows_generator include_meta_data=false
      path = if @sheetfile.start_with? "/xl/" or @sheetfile.start_with? "xl/" then @sheetfile else "xl/#{@sheetfile}" end
      if @book.files.file.exist?(path)
        # SAX parsing, Each element in the stream comes through as two events:
        # one to open the element and one to close it.
        opener = Nokogiri::XML::Reader::TYPE_ELEMENT
        closer = Nokogiri::XML::Reader::TYPE_END_ELEMENT
        Enumerator.new do |y|
          row, cells, cell = nil, {}, nil
          row_number = nil
          cell_type  = nil
          cell_style_idx = nil
          @book.files.file.open(path) do |xml|
            Nokogiri::XML::Reader.from_io(xml).each do |node|
              node_name = node.name
              next unless node_name == CELL || node_name == ROW || node_name == VALUE || node_name == TEXT
              if node_name == ROW
                case node.node_type
                when opener then
                  row = node.attributes
                  row_number = row[ROW_NUMBER]
                  if spans = row['spans']
                    spans = spans.split(":").last.to_i - 1
                  else
                    spans = 0
                  end
                  cells = Array.new(spans)
                  row['cells'] = cells
                  y << (include_meta_data ? row : cells) if node.self_closing?
                when closer
                  y << (include_meta_data ? row : cells)
                end
              elsif (node_name == CELL) && node.node_type == opener
                attributes = node.attributes
                cell_type      = attributes[CELL_TYPE]
                cell_style_idx = attributes[STYLE_INDEX]
                cell           = attributes[CELL_REF]
              elsif node_name == VALUE && node.node_type == opener
                if cell
                  cells[@excel_col_names[cell.sub(row_number, '')]] = convert(node.inner_xml, cell_type, cell_style_idx)
                end
              elsif node_name == TEXT && node.node_type == opener
                if cell
                  cells[@excel_col_names[cell.sub(row_number, '')]] = convert(node.inner_xml, cell_type, cell_style_idx)
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
