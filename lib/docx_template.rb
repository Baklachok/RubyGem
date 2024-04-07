# frozen_string_literal: true

require_relative "docx_template/version"
require "docx"

module DocxTemplate
  class Error < StandardError; end

  class BaseReplacements
    attr_accessor :work_title, :study_direction, :department_name, :city, :year

    def initialize
      # @faculty = 'Факультет'
      @work_title = 'Название работы'
      # @course_number = '№ курса'
      # @group_number = '№ группы'
      # @student = 'Ученик'
      @study_direction = 'Название направления подготовки'
      # @advisor_degree = 'Степень научного руководителя '
      @department_name = 'Название кафедры'
      # @advisor_name = 'ФИО научного руководителя'
      # @total_score = 'Суммарный балл'
      @city = 'Город'
      @year = 'Год'
    end
  end

  class CourseReplacements < BaseReplacements
    attr_accessor :faculty, :course_number, :group_number, :student, :advisor_degree, :advisor_name, :total_score

    def initialize
      super
      @faculty = 'Факультет'
      @course_number = '№ курса'
      @group_number = '№ группы'
      @student = 'Ученик'
      @advisor_degree = 'Степень научного руководителя '
      @advisor_name = 'ФИО научного руководителя'
      @total_score = 'Суммарный балл'
    end
  end

  class IndividualReplacements < BaseReplacements
    attr_accessor :topic

    def initialize
      super
      @topic = 'Тема работы'
    end
  end

  # Your code goes here...
  def self.start
    puts "Please choose option
    1)course_work
    2)graduate_work
    3)individual_work"
    option = gets.chomp.to_i
    create_template(option)
  end

  def self.replace(filename, name, replacements)
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
    true
  end

  def self.create_template(option)
    case option
    when 1
      replacementsClass = CourseReplacements.new
      replacements = {
      "$1": replacementsClass.faculty,
      "$2": replacementsClass.work_title,
      "$3": replacementsClass.course_number,
      "$4": replacementsClass.group_number,
      "$5": replacementsClass.student,
      "$6": replacementsClass.study_direction,
      "$7": replacementsClass.advisor_degree,
      "$8": replacementsClass.department_name,
      "$9": replacementsClass.advisor_name,
      "#0": replacementsClass.total_score,
      "#1": replacementsClass.city,
      "#2": replacementsClass.year,
    }
      replace("./template_course_work.docx", "course_work.docx", replacements)
    when 2
      replace("./template_graduate_work.docx", "graduate_work.docx")
    when 3
      replacementsClass = IndividualReplacements.new
      replacements = {
      "$2": replacementsClass.work_title,
      "$6": replacementsClass.study_direction,
      "$8": replacementsClass.department_name,
      "#1": replacementsClass.city,
      "#2": replacementsClass.year,
      '#3': replacementsClass.topic
    }
      replace("./template_individual_work.docx", "individual_work.docx", replacements)
    end
  end
end
