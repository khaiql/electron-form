require 'axlsx'
require 'byebug'
require './common'

class Form
  include Common

  def initialize
    FileUtils.rm(result_path) if File.exists?(result_path)
    @customer_name = "Khai Le"
    @customer_business = ""
    @customer_street = "Bogenbay Batyra 134"
    @customer_city = "Almaty"
    @customer_postcode = "050000"
    @customer_country_code = "KZ"

    @external_item_id_barcode = "IC280000016KZ"
    @external_item_id = "IC280000016KZ"

    @package_items = [
      {description: "T-Shirt of Cotton", quantity: 2, value: 20.00},
      {description: "T-Shirt of Other Textile Materials", quantity: 1, value: 10.00},
    ]
    @total_value = 30.00
  end

  def result_path
    @result_path ||= File.expand_path("../#{form_name}", __FILE__)
  end

  def form_name
    @form_name ||= "#{self.class.name.downcase}_results.xlsx"
  end

end