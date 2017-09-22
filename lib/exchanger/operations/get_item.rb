module Exchanger
  # The GetItem operation gets items from the Exchanger store.
  # 
  # http://msdn.microsoft.com/en-us/library/aa565934.aspx
  class GetItem < Operation
    class Request < Operation::Request
      attr_accessor :item_ids, :base_shape

      # Reset request options to defaults.
      def reset
        @item_ids = []
        @base_shape = :all_properties
      end

      def to_xml
        Nokogiri::XML::Builder.new do |xml|
          xml.send("soap:Envelope", "xmlns:soap" => NS["soap"],"xmlns" => NS["m"], "xmlns:t" => NS["t"]) do
            xml.send("soap:Header") do 
              xml.send("t:RequestServerVersion", "Version" => "Exchange2013","xmlns"=>"http://schemas.microsoft.com/exchange/services/2006/types", "soap:mustUnderstand"=>"0" ) do
              end
            end
             
            xml.send("soap:Body") do
              xml.GetItem do
                xml.ItemShape do
                  xml.send "t:BaseShape", base_shape.to_s.camelize
                  xml[:t].AdditionalProperties do
                    xml.send "t:FieldURI", "FieldURI"=>"item:Subject"
                    xml.send "t:FieldURI", "FieldURI"=>"item:Body"
                  end
                end
                xml.ItemIds do
                  item_ids.each do |item_id|
                    xml.send("t:ItemId", "Id" => item_id)
                  end
                end
              end
            end
          end
        end
      end
    end

    class Response < Operation::Response
      def items
        to_xml.xpath(".//m:Items", NS).children.map do |node|
          item_klass = Exchanger.const_get(node.name)
          item_klass.new_from_xml(node)
        end
      end
    end
  end
end
