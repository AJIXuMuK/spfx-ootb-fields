declare interface IOotbFieldsStrings {
  DateTime:{[key: string]: string};
  SendEmailTo: string;
  StartChatWith: string;
  Contact: string;
  UpdateProfile: string;
}

declare module 'OotbFieldsStrings' {
  const strings: IOotbFieldsStrings;
  export = strings;
}
