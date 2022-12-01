import { Validation, Approver, MultidayCost } from "../../../models/props";
import {
  IRefObject,
  ITextField,
  Label,
  PrimaryButton,
} from "office-ui-fabric-react/lib";

export interface IReqData {
  formKey: string;

  //Status
  employeeLogin: string;
  personnelNo: string;
  status: string;
  stage: string;
  nextApprover?: number;
  requestLog: string;

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
  // benefitToState: string;

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
  mileageEstimation?: string;
  mileageRate?: number;
  mileageAmount?: string;
  lodgingCostPerNight: string;
  lodgingNights: string;
  totalLodgingAmount?: string;
  mealCostPerNight: string;
  mealPerNights: string;
  totalMealAmount?: string;
  carRentalUsed?: string;
  chbxCarRentalNo: boolean;
  chbxCarRentalYes: boolean;
  vehicleRentalCost?: string;
  otherTransportCosts: string;
  specialMarketingActivitiesAmount?: string;
  numberOfTravelers: string;
  costPerTraveler: string;
  totalEstimatedCostOfTrip?: string;

  //Section D
  TravelerName1: string;
  TravelerjobTitle1: string;
  TravelerName2: string;
  TravelerjobTitle2: string;
  TravelerName3: string;
  TravelerjobTitle3: string;

  //Section E
  agencyAccounting: string;
  deputySecretary: string;
  Agency1: string;
  CostCenter1: string;
  Fund1: string;
  GeneralLedger1: string;
  Grant1: string;
  WBSElemenet1: string;
  Agency2: string;
  CostCenter2: string;
  Fund2: string;
  GeneralLedger2: string;
  Grant2: string;
  WBSElemenet2: string;
  Agency3: string;
  CostCenter3: string;
  Fund3: string;
  GeneralLedger3: string;
  Grant3: string;
  WBSElemenet3: string;

  //Section G
  extraNotes: string;

  //Acct Managers
  employeeApproval: Approver;
  sectionHead?: Approver;
  secretary?: Approver;
  undersecretary?: Approver;
  deputyUndersecretary?: Approver;
  budget?: Approver;
  acctmgr1?: Approver;
  acctmgr2?: Approver;
  acctAdmin?: Approver;

  //Everything Below is useless
  // costCenter: string;
  // domicile: string;
  // taNo: string;
  // departureTime: string;
  // returnTime: string;
  // fund: string;
  // dateOfRequest: Date;
  // fYBudget?: string;
  // amtRemainBudget?: string;
  // amtRemainingAfterThis?: string;
  // authBudget?: string;
  // gL: string;
  // sMAGL: string;
  // fySpecialMarketing?: string;
  // fySpecialMarketingamtRemaining?: string;
  // fySpecialMarketingamtRemainingAfterThis?: string;
  // fYBudgetFY2?: string;
  // amtRemainBudgetFY2?: string;
  // amtRemainingAfterThisFY2?: string;
  // authBudgetFY2?: string;
  // fySpecialMarketingFY2?: string;
  // fySpecialMarketingamtRemainingFY2?: string;
  // fySpecialMarketingamtRemainingAfterThisFY2?: string;
  // airTravelAgencyUsed: string;
  // airTravelAgencyUsedJustification: string;
  // airFare: string;
  // vehicleType: string;
  // vehiclePassengers: string;
  // vehicleRentalTypeIsCompact: string;
  // vehicleRentalJustificationChoice: string;
  // vehicleRentalJustificationText: string;
  // limoTaxi: string;
  // limoTaxiFareAmount?: string;
  // tollsAndParking: string;
  // tollsAndParkingAmount?: string;
  // totalTransportationExpense?: string;
  //lodging: MultidayCost[];
  //meals: MultidayCost[];
  // tips: string;
  // tipsAmount?: string;
  // otherExpensePayableTo: string;
  // otherExpensePaymentMethod: string;
  // otherExpenseDueDate: string;
  // otherExpenseAmount?: string;
  // totalEstimatedTravelAmount?: string;
  // specialMarketingActivitiesAmountNotes: string;
  // travelAdvanceDate: string;
  // travelAdvanceAmount?: string;
  // chbxGPSRentalVehicle: boolean;
  // EstimatedCompensatoryTime: string;
  // budgetYear1?: number;
  // budgetYear2?: number;
  //Everything Above is most likely useless, it's here to we don't hit errors
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
