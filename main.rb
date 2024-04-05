
    require 'docx'



gg = {
    'Факультет':  '$1',
    'Название работы':  '$2',
    '№ курса':  '№ курса',
    '№ группы':  '$4',
    'ученик':  '$5',
    'названиенаправленияподготовки':  '$6',
    'степеньнаучногоруководителя':  '$7',
    'названиекафедры':  '$8',
    'ФИОнаучногоруководителя': '$9',
    'Суммарныйбалл': 'Суммарныйбалл',
    'Город': 'Город',
    'Год': 'Год',
}
doc = Docx::Document.open('./template.docx')

readVariable = Docx::Document.open('./forVariable.docx')
readVariable.paragraphs.each do |readStr|

    # puts readStr.to_s.split("-")
    # if(readStr.to_s ) сделать проверку
    replacement_value = readStr.to_s.split("–")
    puts "asdf " + gg[replacement_value[0].strip.to_sym].to_s + " /" + replacement_value[0].strip + "/"
    replacement_value[0] = gg[replacement_value[0].strip.to_sym].to_s
    replacement_value[1] = replacement_value[1].strip
    # puts "/" + replacement_value[0] + "/"
    # puts replacement_value[1]
    doc.paragraphs.each do |p|

        p.each_text_run do |tr|

            tr.substitute(replacement_value[0], replacement_value[1])      
            # puts tr.to_s + " replace on  " + replacement_value[0]
        end
        # puts replacement_value[0]
        # p.text = p.text.sub(replacement_value[0], replacement_value[1])
        
        # p.styles.font = 'Times New Roman'
        # Создаем стиль с нужным шрифтом и размером текста
        # doc.default_paragraph_style
        # puts "replace " + replacement_value[0] + "  p " +   p

    end

    doc.tables.each do |table|
        last_row = table.rows.last
      
        # Substitute text in each cell of this new row
        last_row .cells.each do |cell|
          cell.paragraphs.each do |paragraph|
            paragraph.each_text_run do |text|
                # puts text
              text.substitute(replacement_value[0], replacement_value[1])
            end
          end
        end
      end
end  



doc.save('example.docx')
