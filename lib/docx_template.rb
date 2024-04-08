# frozen_string_literal: true

require_relative "docx_template/version"
require "docx"

module DocxTemplate
  class Error < StandardError; end

  class BaseReplacements
    attr_accessor :work_title, :study_direction, :department_name, :city, :year

    def initialize
      @work_title = ' '
      @study_direction = '01.03.02 Прикладная математика и информатика'
      @department_name = 'Теории упругости'
      @city = 'Ростов-на-Дону'
      @year = Time.now.strftime("%Y")
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
  
      doc.save(name)
      true
    end
  end

  class CourseReplacements < BaseReplacements
    attr_accessor :faculty, :course_number, :group_number, :student, :advisor_degree, :advisor_name, :total_score

    def initialize
      super
      @faculty = 'Мехмат'
      @course_number = '3'
      @group_number = '5'
      @student = 'Пивоваров Дмитрий Юрьевич'
      @advisor_degree = 'Профессор'
      @advisor_name = 'Мнухин Роман Михайлович'
      @total_score = '100'
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
      @group_number = '1'
      @student = 'Иванов Иван Иванович'
      @advisor_degree = 'Профессор'
      @advisor_name = 'Мнухин Роман Михайлович'

      @order_number = '111-K'
      @order_date = Time.now.strftime("%d.%m.%Y")
      @due_date_student = Time.now.strftime("«%d» %m %Y")
      @initial_data = ' '
      @given_data = ' '
      @solve_problem = ' '
      @subject_area = ''
      @objective = ' '
      @approach = ' '
      @metod_optimize = 'И так сойдет'
      @norm_contorol_fio = 'Д.В. Хроменко'
      @head_of_department = 'Белявский Г. И.'
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
      @topic = ' '
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
end


