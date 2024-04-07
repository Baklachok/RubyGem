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
    replacements = {
      'Факультет': "$1",
      'Название работы': "$2",
      '№ курса': "№ курса",
      '№ группы': "$4",
      'ученик': "$5",
      'названиенаправленияподготовки': "$6",
      'степеньнаучногоруководителя': "$7",
      'названиекафедры': "$8",
      'ФИОнаучногоруководителя': "$9",
      'Суммарныйбалл': "Суммарныйбалл",
      'Город': "Город",
      'Год': "Год"
    }

    puts filename
    doc = Docx::Document.open(filename)
    read_variable_doc = Docx::Document.open("./forVariable.docx")

    read_variable_doc.paragraphs.each do |read_str|
      replacement_value = read_str.to_s.split("–")
      replacement_key = replacement_value[0].strip.to_sym
      replacement_value[0] = replacements[replacement_key].to_s
      replacement_value[1] = replacement_value[1].strip

      doc.paragraphs.each do |p|
        p.each_text_run do |tr|
          tr.substitute(replacement_value[0], replacement_value[1])
        end
      end

      doc.tables.each do |table|
        last_row = table.rows.last
        last_row.cells.each do |cell|
          cell.paragraphs.each do |paragraph|
            paragraph.each_text_run do |text|
              text.substitute(replacement_value[0], replacement_value[1])
            end
          end
        end
      end
    end

    doc.save(name)
    return true
  end

  def self.create_template(option)
    case option
    when 1
      replace("./template_course_work.docx", "../course_work.docx")
    when 2
      replace("./template_graduate_work.docx", "../graduate_work.docx")
    when 3
      replace("./template_individual_work.docx", "../individual_work.docx")
    end
  end
end
