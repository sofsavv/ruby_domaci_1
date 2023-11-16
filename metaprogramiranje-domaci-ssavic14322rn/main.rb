require "google_drive"

session = GoogleDrive::Session.from_config("config.json")
ws = session.spreadsheet_by_key("1ZwnNhN4Uj96DklpoDbJylT8tNVakcjoJK3m9cghfoqQ").worksheets[0]

class Column

    

end

class MyTable 

    include Enumerable
    attr_accessor :table, :worksheet, :row_start, :col_start, :headers

    def initialize(worksheet)
        @worksheet = worksheet
        @row_start = nil
        @col_start = nil
        @headers = []
        @table = []
        find_start
        header
        init_table
        col_method
    end

    private def init_table 

        (@row_start..@worksheet.num_rows).each do |row|
            tmp_row = []
            (@col_start..@worksheet.num_cols).each do |col|
                if @worksheet[row, col] == 'subtotal' || @worksheet[row, col] == 'total'
                    tmp_row.clear
                    break
                end
                tmp_row << worksheet[row, col]
            end
            @table << tmp_row unless tmp_row.empty?
        end
    end

    private def header
        (@col_start..@worksheet.num_cols).each do |col|
            @headers << @worksheet[@row_start, col]
        end
    end

    private def find_start
        (1..@worksheet.num_rows).each do |row|
            (1..@worksheet.num_cols).each do |col|
              if @row_start.nil? && @worksheet[row, col] != ''
                @row_start = row
                @col_start = col
              end
            end
          end
    end

    private def col(col_name)
        col_idx = @headers.index(col_name)
        return unless col_idx
        @table.transpose[col_idx]
    end

    def [](col_name)
        col(col_name)
    end

    def []=(idx, data, col_name)
        # puts "idx: #{idx}"
        col_idx = @headers.index(col_name)
        return unless col_idx
        # puts "data: #{data}"
        @table[idx][col_idx] = data
        @worksheet[idx + @row_start, col_idx + @col_start] = data
        # puts "Updated @worksheet[#{worksheet_row}, #{worksheet_col}] with value #{data}"
        @worksheet.save
        
    end

    private def create_col_meth(col_name, meth_name, &block)
        # col_name.instance_eval{
        #     define_method(meth_name, &block)
        # }
        self.class.send(:define_method, meth_name, &block)
    end

    private def col_method
        @headers.each do |col_name|
            col_name.downcase!
            method_name = col_name.split.map(&:capitalize).join
            method_name[0] = method_name[0].downcase!
            # puts method_name
            create_col_meth(col_name, method_name) do
                col(col_name)
            end
        end
    end


    def row(index)
        @table[index]
    end

    def each
        (1..@worksheet.num_rows).each do |row|
            (1..@worksheet.num_cols).each do |col|
              yield @worksheet[row, col] 
            end
        end
    end
end

t = MyTable.new(ws)
puts t["prva kolona"][2]
t["prva kolona"][2]= 99
# t.[]=("prva kolona", 2, 99)
# puts t["prva"][2]
# puts t.row(3)
# puts t.methods
puts t["prva kolona"]
# puts t.prvaKolona

