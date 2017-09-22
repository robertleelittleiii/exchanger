module Exchanger
  class Field
    attr_accessor :name, :options

    def initialize(name, options = {})
      @name = name
      @options = options
    end

    def type
      options[:type] || String
    end

    def tag_name
      (options[:name] || name).to_s.camelize
    end

    # Only for arrays
    def sub_field
      if type.is_a?(Array)
        Field.new(name, {
            :type => type[0],
            :field_uri_namespace => field_uri
          })
      end
    end

    def field_uri_namespace
      options[:field_uri_namespace].to_s
    end

    def field_uri
      if name == :text
        field_uri_namespace
      else
        "#{field_uri_namespace}:#{tag_name}"
      end
    end

    # FieldURI or IndexedFieldURI
    # <t:FieldURI FieldURI="item:Sensitivity"/>
    # <t:IndexedFieldURI FieldURI="contacts:EmailAddress" FieldIndex="EmailAddress1"/>
    # 
    # http://msdn.microsoft.com/en-us/library/aa494315.aspx
    # http://msdn.microsoft.com/en-us/library/aa581079.aspx
    def to_xml_field_uri(value)
      doc = Nokogiri::XML::Document.new
      if value.is_a?(Entry)
        doc.create_element("IndexedFieldURI", "FieldURI" => field_uri, "FieldIndex" => value.key)
      else
        doc.create_element("FieldURI", "FieldURI" => field_uri)
      end
    end

    # See Element#to_xml_updates.
    # Yields blocks with FieldURI and Item/Folder/etc changes.
    def to_xml_updates(value)
      return if options[:readonly]
      doc = Nokogiri::XML::Document.new
#      puts("* ( ) " * 10)
#      puts(value.class)
#      puts(value.inspect)
#      if  value.is_a?(Array) then 
#        puts(value[0].class)
#        puts(value[0].inspect)
#      end
#      puts("* ( ) " * 10)
#      
#      
      if value.is_a?(Array)
        value.each do |sub_value|
#          if false and sub_value.is_a?(Exchanger::CategoryString) then
#            puts("- * & " * 10)
#            puts("found category string")
#            puts(sub_value.inspect)
#            puts(sub_value.tag_name)
#            puts(sub_value.text)
#            puts("- * & " * 10)
#            element_wrapper=""
#            category_array = sub_value.text.split(",") rescue []
#            category_array.each do |array_item|
#              element_wrapper = doc.create_element(sub_value.tag_name)
#              puts(element_wrapper.to_xml)
#              element_wrapper << array_item
#            end
#            
#            puts(doc.to_xml)
#            puts(element_wrapper)
#
#            yield [
#              field_uri,
#              element_wrapper
#            ]
#            
#          else
            sub_field.to_xml_updates(sub_value) do |field_uri_xml, element_xml|
              element_wrapper = doc.create_element(sub_field.tag_name)
              element_wrapper << element_xml
              yield [
                field_uri_xml,
                element_wrapper
              ]
            end
          end
#        end
      elsif value.is_a?(Exchanger::CategoryString)  #  Categories doesn't follow any other structure so we will create an exception
#        puts("catetory string")
#        puts(value.text)
#        puts(value.text.class)
#        puts(value.inspect)
#        sub_value_array = eval(value.text) rescue []
#        puts(sub_value_array.inspect)
        
        # create a category structure with an array of strings.
        new_node = doc.create_element("Categories")
        value.text.each do |array_element|
          new_node.add_child(doc.create_element("String",array_element))
        end
        
#        puts(self.to_xml_field_uri(value).inspect)
        
        #Create the field URI manually for categories
        
        new_field_uri = doc.create_element("FieldURI")
        new_field_uri["FieldURI"] = "item:Categories"
        
        yield [
          new_field_uri,
          new_node.children
        ]
        
        
        #        sub_field.to_xml_updates(sub_value_array) do |field_uri_xml, element_xml|
        #              element_wrapper = doc.create_element(sub_field.tag_name)
        #              element_wrapper << element_xml
        #              yield [
        #                field_uri_xml,
        #                element_wrapper
        #              ]
        #            end
        #            
        #         yield [
        #          self.to_xml_field_uri(value),
        #          self.to_xml(value)
        #        ]
      elsif value.is_a?(Exchanger::Element)
        value.tag_name = tag_name
        if value.class.elements.keys.include?(:text)
          field = value.class.elements[:text]
          yield [
            field.to_xml_field_uri(value),
            field.to_xml(value)
          ]
        else
          # PhysicalAddress ?
          value.class.elements.each do |name, field|
            yield [
              field.to_xml_field_uri(value),
              field.to_xml(value, :only => [name])
            ]
          end
        end
      else # String, Integer, Boolean, ...
        yield [
          self.to_xml_field_uri(value),
          self.to_xml(value)
        ]
      end
    end

    # Convert Ruby value to XML
    def to_xml(value, options = {})
      if value.is_a?(Exchanger::Element)
        value.tag_name = tag_name
        value.to_xml(options)
      else
        doc = Nokogiri::XML::Document.new
        root = doc.create_element(tag_name)
        case value
        when Array
          value.each do |sub_value|
            root << sub_field.to_xml(sub_value, options)
          end
        when Boolean
          root << doc.create_text_node(value == true)
        when Time
          root << doc.create_text_node(value.xmlschema)
        else # String, Integer, etc
          root << doc.create_text_node(value.to_s)
        end
        root
      end
    end

    # Convert XML to Ruby value
    def value_from_xml(node)
      if type.respond_to?(:new_from_xml)
        type.new_from_xml(node)
      elsif type.is_a?(Array)
        node.children.map do |sub_node|
          sub_field.value_from_xml(sub_node)
        end
      elsif type == Boolean
        node.text == "true"
      elsif type == Integer
        node.text.to_i unless node.text.empty?
      elsif type == Time
        Time.xmlschema(node.text) unless node.text.empty?
      else
        node.text
      end
    end
  end
end
