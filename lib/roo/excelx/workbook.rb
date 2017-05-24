require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Workbook < Excelx::Extractor
      class Label
        attr_reader :sheet, :row, :col, :name

        def initialize(name, sheet, row, col)
          @name = name
          @sheet = sheet
          @row = row.to_i
          @col = ::Roo::Utils.letter_to_number(col)
        end

        def key
          [@row, @col]
        end
      end

      def initialize(path)
        super
        fail ArgumentError, 'missing required workbook file' unless doc_exists?
      end

      def sheets
        doc.xpath('//sheet')
      end

      # aka labels
      def defined_names
        doc.xpath('//definedName').each_with_object({}) do |defined_name, hash|
          # "Sheet1!$C$5"
          sheet, coordinates = defined_name.text.split('!$', 2)
          col, row = coordinates.split('$')
          name = defined_name['name']
          hash[name] = Label.new(name, sheet, row, col)
        end
      end

      def base_timestamp
        @base_timestamp ||= base_date.to_datetime.to_time.to_i
      end

      EPOCH_1900 = Date.new(1900, 1, 1).freeze
      EPOCH_1904 = Date.new(1904, 1, 1).freeze

      def base_date
        @base_date ||=
        begin
          # Default to 1900 (minus one day due to excel quirk) but use 1904 if
          # it's set in the Workbook's workbookPr
          # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx
          result = EPOCH_1900 # default
          doc.css('workbookPr[date1904]').each do |workbookPr|
            if workbookPr['date1904'] =~ /true|1/i
              result = EPOCH_1904
              break
            end
          end

          result
        end
      end
    end
  end
end
