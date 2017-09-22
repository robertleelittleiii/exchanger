module Exchanger
  class Persona < Element
    self.field_uri_namespace = :persona
    self.identifier_name = :persona_id


    #element :mime_content
    element :persona_id, :type => Identifier
    element :parent_folder_id, :type => Identifier
#    element :item_class
#    element :subject
#    element :sensitivity
#    element :body, type: Body
#    element :attachments, :type => [Attachment]
#    element :date_time_received, :type => Time
#    element :size, :type => Integer
#    element :categories, :type => [CategoryString]
#    element :importance
#    element :in_reply_to
#    element :is_submitted, :type => Boolean
#    element :is_draft, :type => Boolean
#    element :is_from_me, :type => Boolean
#    element :is_resend, :type => Boolean
#    element :is_unmodified, :type => Boolean
#    element :internet_message_headers, :type => [String]
#    element :date_time_sent, :type => Time
#    element :date_time_created, :type => Time
#    element :response_objects, :type => [String]
#    element :reminder_due_by, :type => Time
#    element :reminder_is_set, :type => Boolean
#    element :reminder_minutes_before_start
#    element :display_cc
#    element :display_to
#    element :has_attachments, :type => Boolean
#    element :extended_property
#    element :culture
#    element :effective_rights
#    element :last_modified_name
#    element :last_modified_time, :type => Time

#        element :persona_id #=> Identifier
    
#    
#    <PersonaId/>
#   <PersonaType/>
#   <PersonaObjectStatus/>
#   <CreationTime/>
#   <Bodies/>
#   <DisplayNameFirstLastSortKey/>
#   <DisplayNameLastFirstSortKey/>
#   <CompanyNameSortKey/>
#   <HomeCitySortKey/>
#   <WorkCitySortKey/>
#   <DisplayNameFirstLastHeader/>
#   <DisplayNameLastFirstHeader/>
#   <FileAsHeader/>
#   <DisplayName/>
#   <DisplayNameFirstLast/>
#   <DisplayNameLastFirst/>
#   <FileAs/>
#   <FileAsId/>
#   <DisplayNamePrefix/>
element :given_name #   <GivenName/>
element :middle_name#   <MiddleName/>
element :surname #   <Surname/>
#   <Generation/>
element :nickname #   <Nickname/>
#   <YomiCompanyName/>
#   <YomiFirstName/>
#   <YomiLastName/>
element :title#   <Title/>
element :department#   <Department/>
element :company_name#   <CompanyName/>
element :location#   <Location/>
element :email_address, :type=> EmailAddress # <EmailAddress/>
element :email_addresses, :type=> [EmailAddress] #   <EmailAddresses/>
element :phone_number #   <PhoneNumber/>
element :im_address , :type=> ImAddress#   <ImAddress/>
element :home_city #   <HomeCity/>
element :work_city #   <WorkCity/>
#   <RelevanceScore/>
#   <FolderIds/>
#   <Attributions/>
element :display_names, :type=>[String] #   <DisplayNames/>
#   <FileAses/>
#   <FileAsIds/>
#   <DisplayNamePrefixes/>
#   <GivenNames/>
#   <MiddleNames/>
#   <Surnames/>
#   <Generations/>
#   <Nicknames/>
#   <Initials/>
#   <YomiCompanyNames/>
#   <YomiFirstNames/>
#   <YomiLastNames/>
    element :business_phone_numbers, :type => [PhoneNumberPersona] #   <BusinessPhoneNumbers/>
    element :business_phone_numbers2, :type => [PhoneNumberPersona] #   <BusinessPhoneNumbers2/>
    element :home_phones, :type => [PhoneNumberPersona] #   <HomePhones/>
    element :home_phones2, :type => [PhoneNumberPersona] #   <HomePhones2/>
    element :mobile_phones, :type => [PhoneNumberPersona] #   <MobilePhones/>
    element :mobile_phones2, :type => [PhoneNumberPersona] #   <MobilePhones2/>
#   <AssistantPhoneNumbers/>
#   <CallbackPhones/>
#   <CarPhones/>
#   <HomeFaxes/>
#   <OrganizationMainPhones/>
#   <OtherFaxes/>
#   <OtherTelephones/>
#   <OtherPhones2/>
#   <Pagers/>
#   <RadioPhones/>
#   <TelexNumbers/>
#   <TTYTDDPhoneNumbers/>
#   <WorkFaxes/>
#   <Emails1/>
#   <Emails2/>
#   <Emails3/>
#   <BusinessHomePages/>
#   <PersonalHomePages/>
#   <OfficeLocations/>
#   <ImAddresses/>
#   <ImAddresses2/>
#   <ImAddresses3/>
    element :business_addresses, :type => [PhysicalAddress] #   <BusinessAddresses/>
    element :home_addresses, :type => [PhysicalAddress] #   <HomeAddresses/>
    element :other_addresses, :type => [PhysicalAddress]#   <OtherAddresses/>
    element :titles, :type=>[String] #   <Titles/>
    element :departments, :type=>[String] #   <Departments/>
    element :company_names,  :type=>[String] #   <CompanyNames/>
#   <Managers/>
#   <AssistantNames/>
#   <Professions/>
    element :spouse_names, :type=>[String] #   <SpouseNames/>
#   <Children/>
#   <Schools/>
    element :hobbies, :type=>[String] #   <Hobbies/>
    element :wedding_anniversaries, :type=>[Time] #   <WeddingAnniversaries/>
    element :birthdays, :type=>[Time] #   <Birthdays/>
#   <Locations/>
#   <ExtendedProperties/>
    
    
    
    
    element :persona_type
    element :display_name
    element :display_name_first_last
    element :display_name_last_first
    element :email_address, :type => [EmailAddress]
    element :email_addresses, :type => [EmailAddress]
    element :im_address, :type => [ImAddress]
    element :relevance_score
#    element :parent_folder_id, :type => Identifier
#      ["PersonaId", "PersonaType", "DisplayName", "DisplayNameFirstLast", "DisplayNameLastFirst", "EmailAddress", "EmailAddresses", "ImAddress", "RelevanceScore"] 

    def self.find(id)
      find_all([id]).first
    end

    def self.find_all(ids)
      response = GetItem.run(:item_ids => ids)
      response.items
    end

    def self.find_all_by_folder_id(folder_id, max_entries="5000")
      response = FindPeople.run(:folder_id => folder_id, :max_entries= => max_entries)
      response.items
    end

    def self.find_calendar_view_set_by_folder_id(folder_id, calendar_view)
      response = FindItem.run(
        folder_id:     folder_id,
        calendar_view: calendar_view,
      )

      response.items
    end

    attr_writer :parent_folder

    def parent_folder
      @parent_folder ||= if parent_folder_id
        Folder.find(parent_folder_id.id)
      end
    end

    def new_file_attachment(attributes = {})
      FileAttachment.new(attributes.merge(parent_item_id: item_id.id))
    end

    def file_attachments
      attachments.select { |attachment| attachment.is_a?(FileAttachment) }
    end

    def categories_with_color
      parent_folder.category_list.select { |c| categories.include?(c.name) }
    end

    private

    def create
      if parent_folder_id
        options = { folder_id: parent_folder_id.id, items: [self] }.merge(create_additional_options)
        response = CreateItem.run(options)
        self.item_id = response.item_ids[0]
        move_changes
        true
      else
        errors << "Parent folder can't be blank"
        false
      end
    end

    def create_additional_options
      {}  # Implement in subclasses to add CreateItem options
    end

    def update
      if changed?
        options = { items: [self] }.merge(update_additional_options)
        response = UpdateItem.run(options)
        move_changes
        true
      else
        true
      end
    end

    def update_additional_options
      {}  # Implement in subclasses to add UpdateItem options
    end

    def delete
      options = { item_ids: [id] }.merge(delete_additional_options)
      if DeleteItem.run(options)
        true
      else
        false
      end
    end

    def delete_additional_options
      {}  # Implement in subclasses to add DeleteItem options
    end
  end
end
