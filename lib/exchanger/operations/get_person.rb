module Exchanger
  # The GetItem operation gets items from the Exchanger store.
  # 
  # https://msdn.microsoft.com/en-us/library/jj191408(v=exchg.150).aspx
  #
  class GetPerson < Operation
    class Request < Operation::Request
      attr_accessor :person_id, :folder_id, :email_address

      # Reset request options to defaults.
      def reset
        @person_id = ""
        @folder_id = ""
        @email_address
      end

      def to_xml
        Nokogiri::XML::Builder.new do |xml|
          xml.send("soap:Envelope", "xmlns:soap" => NS["soap"],"xmlns:m" => NS["m"], "xmlns:t" => NS["t"]) do
            xml.send("soap:Header") do 
              xml.send("t:RequestServerVersion", "Version" => "Exchange2013","xmlns"=>"http://schemas.microsoft.com/exchange/services/2006/types", "soap:mustUnderstand"=>"0" ) do
              end
            end
             
            xml.send("soap:Body") do
             
              if !folder_id.blank? then
                xml.ParentFolderIds do
                  if folder_id.is_a?(Symbol)
                    xml.send("t:DistinguishedFolderId", "Id" => folder_id) do
                      if email_address
                        xml.send("t:Mailbox") do
                          xml.send("t:EmailAddress", email_address)
                        end
                      end
                    end
                  else
                    xml.send("t:FolderId", "Id" => folder_id)
                  end
                end
              end
              xml[:m].GetPersona do
              
                xml.send("PersonaId", "Id" => person_id)
              
              end
            end
          end
        end
      end
    end

    class Response < Operation::Response
      def elements
        node = to_xml.xpath(".//m:Persona", NS)[0]
        item_klass = Exchanger.const_get(node.name)
        item_klass.new_from_xml(node)
      end
    end
  end
end
