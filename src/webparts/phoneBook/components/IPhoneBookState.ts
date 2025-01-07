export interface IPhoneBookState {
  IsAutocompleteOptionsOpen: boolean;
  SearchValue: string;
  DisplayOptions: Array<any>;
  ContactsOptions: Array<any>;
  IsShowNewRowButton: boolean;
  CurrentPage: number;
  NumOfPages: number;
  PageSize: number;
  PageNumber: number;
  CurrPageDisplayOptions: Array<any>;
  Contacts: Array<any>;
  isLoading: boolean;
  errorMsgDisplay: boolean;
  imgSkeletonFlag: boolean;
  isModalOpen: boolean;
  newUser: {
    displayName: string;
    firstName: string;
    lastName: string;
    phoneNumber: string;
    Email: string;
  };
  errors: {
    firstName?: string;
    lastName?: string;
    Email?: string;
    phoneNumber?: string;
  };
  titleAndUrl: Array<any>;

}
