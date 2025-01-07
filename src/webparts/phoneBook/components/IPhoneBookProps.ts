import { SPFI } from "@pnp/sp";

export interface IPhoneBookProps {
  Contacts: Array<any>;
  context: any;
  PhoneBookTableId: string;
  companyTitle: string;
  sp: SPFI
}
