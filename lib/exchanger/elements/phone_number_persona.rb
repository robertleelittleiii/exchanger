module Exchanger
  # The Entry element represents a telephone number for a contact.
  # 
  # http://msdn.microsoft.com/en-us/library/aa565941.aspx
  class PhoneNumberPersona < Entry
    self.field_uri_namespace = :"personas:PhoneNumber"

    key :key # BusinessFax, BusinessPhone, BusinessPhone2, MobilePhone ...

    # A text value that represents the entry is required if this element is used.
    element :number
    element :type
  end
end
