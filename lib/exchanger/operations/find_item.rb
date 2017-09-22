module Exchanger
  # The FindItem operation identifies items that are located in a specified folder.
  # 
  # http://msdn.microsoft.com/en-us/library/aa566107.aspx
  class FindItem < Operation
    class Request < Operation::Request
      attr_accessor :folder_id, :traversal, :base_shape, :email_address, :calendar_view, :query_string, :offset, :base_point, :max_entries

      # Reset request options to defaults.
      def reset
        @folder_id = :contacts
        @traversal = :shallow
        @base_shape = :all_properties
        @email_address = nil
        @calendar_view = nil
        @query_string = ""
        @max_entries = 1000
        @offset = 0
        @base_point = "Beginning"
  #    <m:IndexedPageItemView MaxEntriesReturned="6" Offset="0" BasePoint="Beginning" />

      end

      def to_xml
        Nokogiri::XML::Builder.new do |xml|
          xml.send("soap:Envelope", "xmlns:soap" => NS["soap"],"xmlns" => NS["m"], "xmlns:t" => NS["t"]) do
            xml.send("soap:Header") do 
              xml.send("t:RequestServerVersion", "Version" => "Exchange2013","xmlns"=>"http://schemas.microsoft.com/exchange/services/2006/types", "soap:mustUnderstand"=>"0" ) do
              end
            end
            xml.send("soap:Body") do
              xml.FindItem( "Traversal" => traversal.to_s.camelize) do
                xml.ItemShape do
                  xml.send "t:BaseShape", base_shape.to_s.camelize
                end
                if calendar_view
                  xml.CalendarView(calendar_view.to_xml.attributes)
                end
                xml.IndexedPageItemView("BasePoint"=>"Beginning", "MaxEntriesReturned"=>"#{@max_entries}" ,"Offset"=>"#{@offset}")

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
                 if !@query_string.blank? then
                  xml.QueryString do
                    xml.text(@query_string)
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
        to_xml.xpath(".//t:Items", NS).children.map do |node|
          item_klass = Exchanger.const_get(node.name)
          item_klass.new_from_xml(node)
        end
      end
      
      def offset
        dataHash = Hash.from_xml(body)
        dataHash["Envelope"]["Body"]["FindItemResponse"]["ResponseMessages"]["FindItemResponseMessage"]["RootFolder"]["IndexedPagingOffset"].to_i || "n'a"
      end
      
      def total_items
        dataHash = Hash.from_xml(body)
        dataHash["Envelope"]["Body"]["FindItemResponse"]["ResponseMessages"]["FindItemResponseMessage"]["RootFolder"]["TotalItemsInView"].to_i || "n'a"
      end
      
    end
  end
end
