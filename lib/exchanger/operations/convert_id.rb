module Exchanger
  # The GetItem operation gets items from the Exchanger store.
  # 
  # https://msdn.microsoft.com/en-us/library/office/bb799665(v=exchg.150).aspx
  
  class ConvertId < Operation
    class Request < Operation::Request
      attr_accessor :item_id, :in_format, :out_format, :mail_box

      # Reset request options to defaults.
      def reset
        @item_id = []
        @in_format = "EwsLegacyId"
        @out_format = "EwsId"
        @mail_box = ""
      end

      def to_xml
        Nokogiri::XML::Builder.new do |xml|
          xml.send("soap:Envelope", "xmlns:soap" => NS["soap"],  "xmlns:t"=>"http://schemas.microsoft.com/exchange/services/2006/types") do
            xml.send("soap:Header") do 
              xml.send("t:RequestServerVersion", "Version" => "Exchange2013" ) do
              end
            end
            xml.send("soap:Body") do
              xml.ConvertId("xmlns" => NS["m"], "xmlns:t" => NS["t"], "DestinationFormat" => @out_format) do
                xml.SourceIds do
                 # xml.send("t:AlternateId", "Format"=>@in_format, "Id"=> @item_id,  "Mailbox"=>@mail_box)
                  xml.send("t:AlternatePublicFolderId", "Format"=>@in_format, "FolderId"=> @item_id)
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
