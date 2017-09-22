module Exchanger
  # The FindItem operation identifies items that are located in a specified folder.
  # 
  # https://msdn.microsoft.com/en-us/library/office/jj191039(v=exchg.150).aspx
  class FindPeople < Operation
    class Request < Operation::Request
      attr_accessor :folder_id, :traversal, :base_shape, :max_entries, :offset, :query_string, :detail_level

      # Reset request options to defaults.
      def reset
        @folder_id = :contacts
        @traversal = :shallow
        @base_shape = :id_only #:all_properties
        @max_entries = 100
        @offset = 0
        @query_string = ""
        @detail_level = :basic #:basic - just name and title, :full  - all info available
      end

      def to_xml
        Nokogiri::XML::Builder.new do |xml|
          xml.send("soap:Envelope", "xmlns:soap" => NS["soap"],  "xmlns:t"=>"http://schemas.microsoft.com/exchange/services/2006/types" , "xmlns:m"=>"http://schemas.microsoft.com/exchange/services/2006/messages") do
            xml.send("soap:Header") do 
              xml.send("t:RequestServerVersion", "Version" => "Exchange2013","xmlns"=>"http://schemas.microsoft.com/exchange/services/2006/types", "soap:mustUnderstand"=>"0" ) do
              end
            end
            xml.send("soap:Body") do
              xml["m"].FindPeople do
                xml["m"].PersonaShape do
                  xml.send "t:BaseShape", base_shape.to_s.camelize
                  xml["t"].AdditionalProperties do
                    if @detail_level == :basic then
                    xml.send("t:FieldURI", "FieldURI"=>"persona:DisplayName")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Title")
                    else
                    xml.send("t:FieldURI", "FieldURI"=>"persona:PersonaId")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:PersonaType")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:PersonaObjectStatus")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:CreationTime")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Bodies")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:DisplayName")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:FileAs")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:FileAsId")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:GivenName")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:MiddleName")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Surname")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:Generation")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:Nickname")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:YomiCompanyName")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:YomiFirstName")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:YomiLastName")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Title")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:Department")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:CompanyName")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:Location")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:EmailAddress")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:EmailAddresses")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:PhoneNumber")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:ImAddress")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:HomeCity")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:WorkCity")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:DisplayName")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:FileAses")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:FileAsIds")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:GivenNames")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:MiddleNames")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Surnames")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Nicknames")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Initials")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:YomiCompanyNames")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:YomiLastNames")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:YomiFirstNames")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:BusinessPhoneNumbers")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:BusinessPhoneNumbers2")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:HomePhones")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:HomePhones2")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:MobilePhones")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:MobilePhones2")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:AssistantPhoneNumbers")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:CarPhones")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:HomeFaxes")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:OrganizationMainPhones")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:OtherFaxes")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:OtherTelephones")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:OtherPhones2")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Pagers")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:RadioPhones")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:TelexNumbers")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:TTYTDDPhoneNumbers")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:WorkFaxes")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Emails1")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Emails2")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Emails3")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:BusinessHomePages")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:PersonalHomePages")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:OfficeLocations")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:ImAddresses1")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:ImAddresses2")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:ImAddresses3")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:BusinessAddresses")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:HomeAddresses")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:OtherAddresses")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Titles")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Departments")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:CompanyNames")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Managers")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:AssistantNames")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Children")
#                    xml.send("t:FieldURI", "FieldURI"=>"persona:Schools")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Hobbies")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:WeddingAnniversaries")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Birthdays")
                    xml.send("t:FieldURI", "FieldURI"=>"persona:Locations")

# XX <PersonaId/>
# XX  <PersonaType/>
# xx  <PersonaObjectStatus/>
# XX  <CreationTime/>
# XX  <Bodies/>
#   <DisplayNameFirstLastSortKey/>
#   <DisplayNameLastFirstSortKey/>
#   <CompanyNameSortKey/>
#   <HomeCitySortKey/>
#   <WorkCitySortKey/>
#   <DisplayNameFirstLastHeader/>
#   <DisplayNameLastFirstHeader/>
# XX  <FileAsHeader/>
# XX  <DisplayName/>
# XX  <DisplayNameFirstLast/>
# XX  <DisplayNameLastFirst/>
# XX  <FileAs/>
# XX  <FileAsId/>
#   <DisplayNamePrefix/>
# XX   <GivenName/>
# XX  <MiddleName/>
# XX  <Surname/>
# XX  <Generation/>
# XX  <Nickname/>
# XX  <YomiCompanyName/>
# XX  <YomiFirstName/>
# XX  <YomiLastName/>
# XX  <Title/>
# XX  <Department/>
# XX  <CompanyName/>
# XX  <Location/>
# XX  <EmailAddress/>
# XX  <EmailAddresses/>
# XX  <PhoneNumber/>
# XX  <ImAddress/>
# XX  <HomeCity/>
# XX  <WorkCity/>
#   <RelevanceScore/>
#   <FolderIds/>
#   <Attributions/>
# XX  <DisplayNames/>
# XX  <FileAses/>
# XX  <FileAsIds/>
# XX  <DisplayNamePrefixes/>
# XX  <GivenNames/>
# XX  <MiddleNames/>
# XX  <Surnames/>
# XX  <Generations/>
# XX  <Nicknames/>
# XX  <Initials/>
# XX  <YomiCompanyNames/>
# XX  <YomiFirstNames/>
# XX  <YomiLastNames/>
# XX  <BusinessPhoneNumbers/>
# XX  <BusinessPhoneNumbers2/>
# XX  <HomePhones/>
# XX  <HomePhones2/>
# XX  <MobilePhones/>
# xx  <MobilePhones2/>
# XX  <AssistantPhoneNumbers/>
# XX  <CallbackPhones/>
# XX  <CarPhones/>
# XX  <HomeFaxes/>
# XX  <OrganizationMainPhones/>
# XX  <OtherFaxes/>
# XX  <OtherTelephones/>
# XX  <OtherPhones2/>
# XX  <Pagers/>
# XX  <RadioPhones/>
# XX  <TelexNumbers/>
# XX  <TTYTDDPhoneNumbers/>
# XX  <WorkFaxes/>
# XX  <Emails1/>
# XX  <Emails2/>
# XX  <Emails3/>
# XX  <BusinessHomePages/>
# XX  <PersonalHomePages/>
# XX  <OfficeLocations/>
# XX  <ImAddresses/>
# XX  <ImAddresses2/>
# XX  <ImAddresses3/>
# XX  <BusinessAddresses/>
# XX  <HomeAddresses/>
# XX  <OtherAddresses/>
# XX  <Titles/>
# XX  <Departments/>
# XX  <CompanyNames/>
# XX  <Managers/>
# XX  <AssistantNames/>
# XX  <Professions/>
# XX  <SpouseNames/>
# XX  <Children/>
# XX  <Schools/>
# XX  <Hobbies/>
# XX  <WeddingAnniversaries/>
# XX  <Birthdays/>
# XX  <Locations/>
# XX  <ExtendedProperties/>
                    end
                    
                  end
                  
                end
                
                xml["m"].IndexedPageItemView("BasePoint"=>"Beginning", "MaxEntriesReturned"=>"#{@max_entries}" ,"Offset"=>"#{@offset}")
                xml["m"].ParentFolderId do
                  if folder_id.is_a?(Symbol)
                    xml.send("t:DistinguishedFolderId", "Id" => folder_id) do
                    end
                  else
                    # xml.send("t:AddressListId", "Id" => folder_id)
                    xml.send("t:FolderId", "Id" => folder_id)
                  end
                end
                if !@query_string.blank? then
                  xml["m"].QueryString do
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
        to_xml.xpath(".//t:Persona", NS).map do |node|
          item_klass = Exchanger.const_get(node.name)
          item_klass.new_from_xml(node)
        end
      end
    end
  end
end
