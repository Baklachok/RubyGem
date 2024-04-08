# DocxTemplate

In this gem, 3 classes of templates are available for creating completed work. Each class has fields responsible for filling data.  

## Installation

$ gem install bundler  
$ bundle init  
Fill Gemfile with:  
gem 'docx_template', :git => 'https://github.com/Baklachok/RubyGem.git'  
$ bundle install  
and use it  

TODO: Replace `UPDATE_WITH_YOUR_GEM_NAME_PRIOR_TO_RELEASE_TO_RUBYGEMS_ORG` with your gem name right after releasing it to RubyGems.org. Please do not do it earlier due to security reasons. Alternatively, replace this section with instructions to install your gem from git if you don't plan to release to RubyGems.org.  

Install the gem and add to the application's Gemfile by executing:  

    $ bundle add UPDATE_WITH_YOUR_GEM_NAME_PRIOR_TO_RELEASE_TO_RUBYGEMS_ORG  

If bundler is not being used to manage dependencies, install the gem by executing:  

    $ gem install UPDATE_WITH_YOUR_GEM_NAME_PRIOR_TO_RELEASE_TO_RUBYGEMS_ORG  

## Usage
All classes have these fields (default values):  
    work_title = ' '  
    study_direction = '01.03.02 Прикладная математика и информатика'  
    department_name = 'Теории упругости'  
    city = 'Ростов-на-Дону'  
    year = Time.now.strftime("%Y")  

Class CourseReplacements fields (default value):  
    faculty = 'Мехмат'  
    course_number = '3'  
    group_number = '5'  
    student = 'Пивоваров Дмитрий Юрьевич'  
    advisor_degree = 'Профессор'  
    advisor_name = 'Мнухин Роман Михайлович'  
    total_score = '100'  

Class IndividualReplacements fields (default value):  
    topic = ' '  

Class GraduateReplacements fields (default value):  
    group_number = '1'  
    student = 'Иванов Иван Иванович'  
    advisor_degree = 'Профессор'  
    advisor_name = 'Мнухин Роман Михайлович'  
    order_number = '111-K'  
    order_date = Time.now.strftime("%d.%m.%Y")  
    due_date_student = Time.now.strftime("«%d» %m %Y")  
    initial_data = ' '  
    given_data = ' '  
    solve_problem = ' '  
    subject_area = ''  
    objective = ' '  
    approach = ' '  
    metod_optimize = 'И так сойдет'  
    norm_contorol_fio = 'Д.В. Хроменко'  
    head_of_department = 'Белявский Г. И.'  

Examples of crating different files  

$ course_work = DocxTemplate::CourseReplacements.new  
$ course_work.faculty = 'Институт Математики, Механики и Компьютерных Наук имени И.И. Воровича'  
$ course_work.create_word_file()  


$ individual_work = DocxTemplate::IndividualReplacements.new  
$ individual_work.year = 2006  
$ individual_work.create_word_file()  

$ graduate_work = DocxTemplate::GraduateReplacements.new  
$ graduate_work.approach = 'Наш подход'  
$ graduate_work.create_word_file()  

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.  

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).  

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/docx_template. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/[USERNAME]/docx_template/blob/main/CODE_OF_CONDUCT.md).  

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the DocxTemplate project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/[USERNAME]/docx_template/blob/main/CODE_OF_CONDUCT.md).
