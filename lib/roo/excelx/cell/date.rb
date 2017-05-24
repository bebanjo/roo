require 'date'

module Roo
  class Excelx
    class Cell
      class Date < Roo::Excelx::Cell::DateTime
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :coordinate

        attr_reader_with_default default_type: :date

        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          # NOTE: Pass all arguments to the parent class, DateTime.
          super
          @format = excelx_type.last
          @value = link ? Roo::Link.new(link, value) : create_date(base_date, value)
        end

        private

        def create_datetime(_,_);  end

        def create_date(base_date, value)
          offset = offset_for_base_date(base_date, value.to_i)
          base_date + offset
        end

        def offset_for_base_date(base_date, offset)
          # Adjust for Excel erroneously treating 1900 as a leap year
          if EPOCH_1900 == base_date
            offset -= 1

            if offset > 58
              offset -= 1
            end
          end

          offset
        end
      end
    end
  end
end
