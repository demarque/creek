require 'zip/filesystem'
require 'nokogiri'

module Creek

  class Creek::SharedStrings

    attr_reader :book, :dictionary

    def initialize book
      @book = book
      @dictionary = parse_shared_shared_strings
    end

    def parse_shared_shared_strings
      path = "xl/sharedStrings.xml"
      if @book.files.file.exist?(path)
        @book.files.file.open(path) do |doc|
          dictionary = []
          buffer = false
          str = nil
          Nokogiri::XML::Reader.from_io(doc).each do |node|
            case node.node_type
            when Nokogiri::XML::Reader::TYPE_ELEMENT
              case node.name
              when 'si' then
                str = ''
              when 't' then
                buffer = true
              end
            when Nokogiri::XML::Reader::TYPE_END_ELEMENT
              case node.name
              when 'si' then
                dictionary << str
                str = nil
              when 't' then
                buffer = false
              end
            when Nokogiri::XML::Reader::TYPE_TEXT
              str << node.value if buffer
            end
          end
          dictionary
        end
      end
    end
  end
end

