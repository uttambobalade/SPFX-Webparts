import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAnonymousApInExtLibProps {
  description: string;
  siteURL: string;
  userId: string;
  context: WebPartContext;
}
