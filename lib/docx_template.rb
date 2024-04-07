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

    def replace(filename, name, replacements)
      doc = Docx::Document.open(filename)

  
        doc.paragraphs.each do |p|

          p.each_text_run do |tr|
            if(replacements.keys.include?(tr.to_s.to_sym))
              tr.substitute(tr.to_s, replacements[tr.to_s.to_sym].to_s)
            end
          end
        end
  
        doc.tables.each do |table|
          last_row = table.rows.last
          last_row.cells.each do |cell|
            cell.paragraphs.each do |paragraph|
              paragraph.each_text_run do |text|
                if(replacements.keys.include?(text.to_s.to_sym))
                  text.substitute(text.to_s, replacements[text.to_s.to_sym].to_s)
                end
              end
            end
          end
        end
      # end
  
      doc.save(name)
      true
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
    
    def create_word_file()
      replacements = {
        "$1": self.faculty,
        "$2": self.work_title,
        "$3": self.course_number,
        "$4": self.group_number,
        "$5": self.student,
        "$6": self.study_direction,
        "$7": self.advisor_degree,
        "$8": self.department_name,
        "$9": self.advisor_name,
        "#0": self.total_score,
        "#1": self.city,
        "#2": self.year,
      }
      replace("#{File.dirname(File.expand_path(__FILE__))}/template_course_work.docx", "./course_work.docx", replacements)

    end
  end

  class GraduateReplacements < BaseReplacements
    attr_accessor :group_number, :student, :advisor_degree, :advisor_name, :order_number, :order_date, :due_date_student, :initial_data , :given_data , :solve_problem, :subject_area,
                  :objective, :approach , :metod_optimize, :norm_contorol_fio, :head_of_department

    def initialize
      super
      @group_number = '№ группы'
      @student = 'Ученик'
      @advisor_degree = 'Степень научного руководителя '
      @advisor_name = 'ФИО научного руководителя'

      @order_number = '111-K'
      @order_date = 'Date'
      @due_date_student = '«20» 02 2024'
      @initial_data = 'Входные данные'
      @given_data = 'ДАНО'
      @solve_problem = 'Solve'
      @subject_area = 'структурная схема информационных потоков строительной компании'
      @objective = 'разработать информационную систему'
      @approach = 'информационные системы'
      @metod_optimize = 'на основе математического моделирования'
      @norm_contorol_fio = 'Д.В. Хроменко'
      @head_of_department = 'Белявский'
    end

    def create_word_file()
      replacements = {
       "$0": self.group_number,
       "$1": self.student,
       "$2": self.advisor_degree,
       "$3": self.advisor_name,
       "$4": self.order_number,
       "$5": self.order_date,
       "$6": self.due_date_student,
       "$7": self.initial_data,
       "$8": self.given_data,
       "$9": self.solve_problem,
       "#0": self.subject_area,
       "#1": self.objective,
       "#2": self.approach,
       "#3": self.metod_optimize,
       "#4": self.norm_contorol_fio,
       "#5": self.head_of_department,
       "#6": self.work_title,
       "#7": self.study_direction,
       "#8": self.department_name,
       "#9": self.city,
       "!0": self.year,
      }
      replace("#{File.dirname(File.expand_path(__FILE__))}/template_graduate_work.docx", "./graduate_work.docx", replacements)
    end
  end


 

  class IndividualReplacements < BaseReplacements
    attr_accessor :topic

    def initialize
      super
      @topic = 'Тема работы'
    end
    
    def create_word_file()
      replacements = {
        "$2": self.work_title,
        "$6": self.study_direction,
        "$8": self.department_name,
        "#1": self.city,
        "#2": self.year,
        '#3': self.topic
      }
        replace("#{File.dirname(File.expand_path(__FILE__))}/template_individual_work.docx", "./individual_work.docx", replacements)
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


      doc.paragraphs.each do |p|

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
      replace("#{File.dirname(File.expand_path(__FILE__))}/template_course_work.docx", "./course_work.docx", replacements)
    when 2
      replacementsClass = GraduateReplacements.new
      replacements = {
       "$0": replacementsClass.group_number,
       "$1": replacementsClass.student,
       "$2": replacementsClass.advisor_degree,
       "$3": replacementsClass.advisor_name,
       "$4": replacementsClass.order_number,
       "$5": replacementsClass.order_date,
       "$6": replacementsClass.due_date_student,
       "$7": replacementsClass.initial_data,
       "$8": replacementsClass.given_data,
       "$9": replacementsClass.solve_problem,
       "#0": replacementsClass.subject_area,
       "#1": replacementsClass.objective,
       "#2": replacementsClass.approach,
       "#3": replacementsClass.metod_optimize,
       "#4": replacementsClass.norm_contorol_fio,
       "#5": replacementsClass.head_of_department,
       "#6": replacementsClass.work_title,
       "#7": replacementsClass.study_direction,
       "#8": replacementsClass.department_name,
       "#9": replacementsClass.city,
       "!0": replacementsClass.year,
      }
      replace("#{File.dirname(File.expand_path(__FILE__))}/template_graduate_work.docx", "./graduate_work.docx", replacements)
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
      replace("#{File.dirname(File.expand_path(__FILE__))}/template_individual_work.docx", "./individual_work.docx", replacements)
    end
  end
end
