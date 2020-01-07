import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IGiftRegisterProps {
  description: string;
  context : WebPartContext;
  itemID : string;
  viewMode : string;
}

export enum Status {
  Submitted = "Submitted",
  Approved = "Approved",
  Approver1Pending = "Awaiting Division Head 1 Approval",
  Approver2Pending = "Awaiting Division Head 2 Approval",
  RejectedManager = "Rejected - Manager",
  RejectedApprover1 = "Rejected - Division Head 1",
  RejectedApprover2 = "Rejected - Division Head 2"
}