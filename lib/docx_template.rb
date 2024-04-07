# frozen_string_literal: true

require_relative "docx_template/version"
require "docx"

module DocxTemplate
  class Error < StandardError; end

  # Your code goes here...
  def self.start
    puts "Please choose option
    1)course_work
    2)graduate_work
    3)individual_work"
    option = gets.chomp.to_i
    create_template(option)
  end

  def self.replace(filename, name)
    # replacements = {
    #   'Факультет': "$A",
    #   'Название работы': "$B",
    #   '№ курса': "$C",
    #   '№ группы': "$D",
    #   'ученик': "$E",
    #   'названиенаправленияподготовки': "$F",
    #   'степеньнаучногоруководителя': "$G",
    #   'названиекафедры': "$H",
    #   'ФИОнаучногоруководителя': "$I",
    #   'Суммарныйбалл': "$J",
    #   'Город': "$K",
    #   'Год': "$L"
    # }
    replacements = {
      "$1": "Отдел По Борьбе С П@нтами и активистами, а также Симигуками",
      "$2": "Сформулировать проблему диверсификации инвестиционных рисков",
      "$3": "666",
      "$4": "228",
      "$5": "Абдулхамидов Арамановолваолва Ибналохмат",
      "$6": "менеджмет",
      "$7": "ееее",
      "$8": "кафедра суетологии",
      "$9": "Абдулхамидов Арамановолваолва Ибналохмат",
      "#0": "777",
      "#1": "Дорог",
      "#2": "2222"
    }

    doc = Docx::Document.open(filename)
    # read_variable_doc = Docx::Document.open("./forVariable.docx")

    # read_variable_doc.paragraphs.each do |read_str|
      # for (key, item) in read_Variable
        
      # end
      # replacement_value = read_str.to_s.split("–")
      # replacement_key = replacement_value[0].strip.to_sym
      # replacement_value[0] = replacements[replacement_key].to_s
      # replacement_value[1] = replacement_value[1].strip

      doc.paragraphs.each do |p|
        # if (p.text.include?("Студента $C курса $D группы"))
        #   p.ass
        # end
        p.each_text_run do |tr|
          if(replacements.keys.include?(tr.to_s.to_sym))
            tr.substitute(tr.to_s, replacements[tr.to_s.to_sym])
          end
        end
      end

      doc.tables.each do |table|
        last_row = table.rows.last
        last_row.cells.each do |cell|
          cell.paragraphs.each do |paragraph|
            paragraph.each_text_run do |text|
              if(replacements.keys.include?(text.to_s.to_sym))
                text.substitute(text.to_s, replacements[text.to_s.to_sym])
              end
              # text.substitute(replacement_value[0], replacement_value[1])
            end
          end
        end
      end
    # end

    doc.save(name)
  end

  def self.create_template(option)
    case option
    when 1
      replace("./template_course_work.docx", "course_work.docx")
    when 2
      replace("./template_graduate_work.docx", "graduate_work.docx")
    when 3
      replace("./template_individual_work.docx", "individual_work.docx")
    end
  end
end
