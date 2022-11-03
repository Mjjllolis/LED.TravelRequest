export class Validation {
  public controlName: string;
  public message: string;
}

export interface RequestEmp {
  Title: string;
  EMail: string;
  JobTitle: string;
  FirstName: string;
  LastName: string;
}

export class Expense {
  public type: string;
  public description: string;
  public amount: number;
  public subtotal: number;
  public subcompactCar: boolean;
  public subcompactReason: string;
  public carOwnership: string;
  public miles: number;
  public centsPerMile: number;
  public explinationOrListOfPassangers: string;
  public isMemberOf: boolean;
}

export interface Approver {
  userLogin: string;
  jobTitle: string;
  displayName: string;
  approvalStatus: string;
  approvalDate: Date;
  comment: string;
  userId: number;
  approvalString?: string;
}

export class MultidayCost {
  public i?: number;
  public days?: number;
  public cost?: number;
  public total?: number;
}

export class AdditionalTravelerClass {
  public TravelerName: string;
  public TravelerjobTitle: string;
}

export class AgencyAccountingClass {
  public Agency: string;
  public CostCenter: string;
  public Fund: string;
  public GeneralLedger: string;
  public Grant: string;
  public WBSElemenet: string;
}
