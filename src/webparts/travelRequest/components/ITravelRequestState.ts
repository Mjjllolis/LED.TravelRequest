import {
  Validation,
  Approver,
  MultidayCost,
  AdditionalTravelerClass,
  AgencyAccountingClass,
} from "../../../models/props";
import {
  IRefObject,
  ITextField,
  Label,
  PrimaryButton,
} from "office-ui-fabric-react/lib";

export interface IReqData {
  formKey: string;

  //Section A
  employeeId: number;
  employeeName: string;
  employeeTitle: string;
  departureDate?: Date;
  departureDateStr: string;
  returnDate?: Date;
  returnDateStr: string;
  destination: string;
  agency: string;
  division: string;
  modeOfTransportation: string;
  justficationForTrip: string;

  //Section B
  chbxConferenceSeminar: boolean;
  chbxAnnualAuthForTravel: boolean;
  chbxInStateTravel: boolean;
  chbxOutOfStateTravel: boolean;
  chbxWeekend: boolean;
  chbxVehicleRental: boolean;
  chbxUserOfPersonalVehicle: boolean;
  chbxSpecialMarketingActivities: boolean;
  chbxProspectInSameHotelAsEmployee: boolean;
  chbx50pctLodgingException: boolean;
  chbxOther: boolean;
  chbxVehicleRentalSig: string;
  chbxGPSRentalVehicleSig: string;
  chbxProspectInSameHotelAsEmployeeSig: string;
  chbxSpecialMarketingActivitiesSig: string;
  chbx50pctLodgingExceptionSig: string;
  chbxOtherSig: string;

  //Section C
  registrationFees: string;
  airFareCost?: string;
  mileageEstimation?: number;
  mileageRate?: number;
  mileageAmount?: string;
  lodgingCostPerNight: string;
  lodgingNights: string;
  totalLodgingAmount?: string;
  mealCostPerNight: string;
  mealPerNights: string;
  totalMealAmount?: string;
  chbxCarRentalYes: boolean;
  chbxCarRentalNo: boolean;
  vehicleRentalCost?: string;
  specialMarketingActivitiesAmount?: string;
  totalEstimatedCostOfTrip?: string;

  //Section D
  //AdditionalTraveler: AdditionalTravelerClass[]; //Might use array later, for now just want it working
  TravelerName: string;
  TravelerjobTitle: string;

  //Section E
  agencyAccounting: string;
  deputySecretary: string;
  //AgencyAccounting: AgencyAccountingClass[]; //Might use array later, for now just want it working
  Agency: string;
  CostCenter: string;
  Fund: string;
  GeneralLedger: string;
  Grant: string;
  WBSElemenet: string;

  //Section G
  extraNotes: string;

  //Everything Below is useless, it's here to we don't hit errors
  employeeLogin: string;
  personnelNo: string;
  costCenter: string;
  domicile: string;
  taNo: string;
  departureTime: string;
  returnTime: string;
  fund: string;
  dateOfRequest: Date;
  fYBudget?: string;
  amtRemainBudget?: string;
  amtRemainingAfterThis?: string;
  authBudget?: string;
  benefitToState: string;
  gL: string;
  sMAGL: string;
  fySpecialMarketing?: string;
  fySpecialMarketingamtRemaining?: string;
  fySpecialMarketingamtRemainingAfterThis?: string;
  fYBudgetFY2?: string;
  amtRemainBudgetFY2?: string;
  amtRemainingAfterThisFY2?: string;
  authBudgetFY2?: string;
  fySpecialMarketingFY2?: string;
  fySpecialMarketingamtRemainingFY2?: string;
  fySpecialMarketingamtRemainingAfterThisFY2?: string;

  status: string;
  stage: string;
  nextApprover?: number;
  requestLog: string;

  airTravelAgencyUsed: string;
  airTravelAgencyUsedJustification: string;
  airFare: string;
  vehicleType: string;
  vehiclePassengers: string;
  vehicleRentalTypeIsCompact: string;
  vehicleRentalJustificationChoice: string;
  vehicleRentalJustificationText: string;
  limoTaxi: string;
  limoTaxiFareAmount?: string;
  tollsAndParking: string;
  tollsAndParkingAmount?: string;
  totalTransportationExpense?: string;
  lodging: MultidayCost[];
  meals: MultidayCost[];
  tips: string;
  tipsAmount?: string;
  otherExpensePayableTo: string;
  otherExpensePaymentMethod: string;
  otherExpenseDueDate: string;
  otherExpenseAmount?: string;
  totalEstimatedTravelAmount?: string;
  specialMarketingActivitiesAmountNotes: string;
  travelAdvanceDate: string;
  travelAdvanceAmount?: string;
  chbxGPSRentalVehicle: boolean;

  EstimatedCompensatoryTime: string;
  employeeApproval: Approver;
  sectionHead?: Approver;
  secretary?: Approver;
  undersecretary?: Approver;
  deputyUndersecretary?: Approver;
  budget?: Approver;
  acctmgr1?: Approver;
  acctmgr2?: Approver;
  acctAdmin?: Approver;
  budgetYear1?: number;
  budgetYear2?: number;
  //Everything Above is useless, it's here to we don't hit errors
}
export interface ITravelRequestState {
  error: boolean;
  message: string;
  results: any[];
  validations: Validation[];
  textInput: IRefObject<ITextField>;
  AddingAttachment: boolean;
  Attachments: any[];
  reqData: IReqData;
  kickoffFLOW: string;

  hideDialog: boolean;
  dialogTitle: string;
  dialogText: string;
  requestID: string;
  formMode: string;
  saving: boolean;
  printing: boolean;
  DepartureDateError?: string;
  ReturnDateError?: string;
}
