import * as React from "react";
import styles from "./TravelRequest.module.scss";
import { ITravelRequestProps } from "./ITravelRequestProps";
import { ITravelRequestState } from "./ITravelRequestState";
import { Validation, Approver, MultidayCost } from "../../../models/props";
import { stringIsNullOrEmpty, getRandomString } from "@pnp/common";
import { WebEnsureUserResult, sp } from "@pnp/sp";
import { escape } from "@microsoft/sp-lodash-subset";
import { DataService } from "../../../services/data-service";
import ReqInput from "../controls/Input";
import {
  TextField,
  Label,
  PrimaryButton,
  DefaultButton,
  Grid,
  Dialog,
  DialogFooter,
  DialogType,
  Toggle,
  ChoiceGroup,
  IChoiceGroupOption,
  Checkbox,
  IDropdownOption,
  Dropdown,
  IStackTokens,
  DatePicker,
  IDatePickerStrings,
  ActionButton,
  IIconProps,
  Stack,
  Spinner,
  MaskedTextField,
} from "office-ui-fabric-react/lib";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import "office-ui-fabric-core/dist/css/fabric.min.css";
import {
  DateTimePicker,
  DateConvention,
  TimeConvention,
} from "@pnp/spfx-controls-react/lib/dateTimePicker";
import * as CurrencyFormat from "react-currency-format";
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";
import AddAttachment from "./AddAttachment";
import { ToastContainer, Toast, toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import MaskedInput from "react-maskedinput";
import CurrencyTextField from "@unicef/material-ui-currency-textfield";

const DatePickerStrings: IDatePickerStrings = {
  months: [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ],
  shortMonths: [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ],
  days: [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ],
  shortDays: ["S", "M", "T", "W", "T", "F", "S"],
  goToToday: "Go to today",
  prevMonthAriaLabel: "Go to previous month",
  nextMonthAriaLabel: "Go to next month",
  prevYearAriaLabel: "Go to previous year",
  nextYearAriaLabel: "Go to next year",
  closeButtonAriaLabel: "Close date picker",
  isRequiredErrorMessage: "Date is required.",
  invalidInputErrorMessage: "Invalid date format.",
};

const DDHours: IDropdownOption[] = [
  { key: "01", text: "01" },
  { key: "02", text: "02" },
  { key: "03", text: "03" },
  { key: "04", text: "04" },
  { key: "05", text: "05" },
  { key: "06", text: "06" },
  { key: "07", text: "07" },
  { key: "08", text: "08" },
  { key: "09", text: "09" },
  { key: "10", text: "10" },
  { key: "11", text: "11" },
  { key: "12", text: "12" },
  { key: "13", text: "13" },
  { key: "14", text: "14" },
  { key: "15", text: "15" },
  { key: "16", text: "16" },
  { key: "17", text: "17" },
  { key: "18", text: "18" },
  { key: "19", text: "19" },
  { key: "20", text: "20" },
  { key: "21", text: "21" },
  { key: "22", text: "22" },
  { key: "23", text: "23" },
  { key: "24", text: "24" },
];

const DDMinutes: IDropdownOption[] = [
  { key: "00", text: "00" },
  { key: "15", text: "15" },
  { key: "30", text: "30" },
  { key: "45", text: "45" },
];

const stackTokens: IStackTokens = { childrenGap: 5 };

const checkboxStyles = () => {
  return {
    root: {
      marginTop: "1px",
    },
  };
};

const specialApproverSigStyles = () => {
  return {
    root: {
      height: "18px",
    },
  };
};

export default class TravelRequest extends React.Component<
  ITravelRequestProps,
  ITravelRequestState
> {
  private service: DataService;

  constructor(props) {
    super(props);

    this.state = {
      error: false,
      message: "",
      results: [],
      validations: [],
      AddingAttachment: false,
      Attachments: [],
      kickoffFLOW: "",

      reqData: {
        formKey: "",
        personnelNo: "",
        status: "Draft",
        stage: "",
        nextApprover: null,
        requestLog: "",
        dateOfRequest: new Date(),

        //Section A
        employeeId: null,
        employeeName: "",
        employeeTitle: "",
        employeeLogin: "",
        agency: "",
        destination: "",
        departureDateStr: "",
        returnDateStr: "",
        division: "",
        justficationForTrip: "",
        modeOfTransportation: "",

        //Section B
        chbxConferenceSeminar: false,
        chbxAnnualAuthForTravel: false,
        chbxInStateTravel: false,
        chbxOutOfStateTravel: false,
        chbxWeekend: false,
        chbxVehicleRental: false,
        chbxUserOfPersonalVehicle: false,
        chbxSpecialMarketingActivities: false,
        chbxProspectInSameHotelAsEmployee: false,
        chbx50pctLodgingException: false,
        chbxProspectInSameHotelAsEmployeeSig: "",
        chbxSpecialMarketingActivitiesSig: "",
        chbx50pctLodgingExceptionSig: "",
        chbxOther: false,
        chbxOtherSig: "",

        //Section C
        registrationFees: "",
        airFareCost: "",
        mileageEstimation: "", //Total miles
        mileageRate: 0.0, //This is defined in form, not sure it's needed
        mileageAmount: "",
        lodgingCostPerNight: "",
        lodgingNights: "",
        totalLodgingAmount: "",
        mealCostPerNight: "",
        mealPerNights: "",
        totalMealAmount: "",
        chbxCarRentalYes: false,
        chbxCarRentalNo: false,
        carRentalUsed: null,
        chbxVehicleRentalSig: "",
        chbxGPSRentalVehicleSig: "",
        vehicleRentalCost: "",
        otherTransportCosts: "",
        numberOfTravelers: "",
        costPerTraveler: "",
        specialMarketingActivitiesAmount: "",
        totalEstimatedCostOfTrip: "",

        //Section D
        TravelerName1: "",
        TravelerjobTitle1: "",
        TravelerName2: "",
        TravelerjobTitle2: "",
        TravelerName3: "",
        TravelerjobTitle3: "",

        //Section E
        agencyAccounting: "",
        deputySecretary: "",
        Agency1: "",
        CostCenter1: "",
        Fund1: "",
        GeneralLedger1: "",
        Grant1: "",
        WBSElemenet1: "",
        Agency2: "",
        CostCenter2: "",
        Fund2: "",
        GeneralLedger2: "",
        Grant2: "",
        WBSElemenet2: "",
        Agency3: "",
        CostCenter3: "",
        Fund3: "",
        GeneralLedger3: "",
        Grant3: "",
        WBSElemenet3: "",

        //Section G
        extraNotes: "",

        //Acct Managers
        employeeApproval: {
          userLogin: "",
          jobTitle: "Employee",
          displayName: "",
          approvalStatus: "",
          approvalDate: new Date(),
          comment: "",
          userId: null,
        },
        sectionHead: {
          userLogin: "",
          jobTitle: "Section Head",
          displayName: "",
          approvalStatus: "",
          approvalDate: null,
          comment: "",
          userId: null,
          approvalString: "",
        },
        secretary: {
          userLogin: "",
          jobTitle: "Secretary",
          displayName: "",
          approvalStatus: "",
          approvalDate: null,
          comment: "",
          userId: null,
          approvalString: "",
        },
        undersecretary: {
          userLogin: "",
          jobTitle: "Undersecretary",
          displayName: "",
          approvalStatus: "",
          approvalDate: null,
          comment: "",
          userId: null,
          approvalString: "",
        },
        deputyUndersecretary: {
          userLogin: "",
          jobTitle: "Deputy Undersecretary",
          displayName: "",
          approvalStatus: "",
          approvalDate: null,
          comment: "",
          userId: null,
          approvalString: "",
        },
        budget: {
          userLogin: "",
          jobTitle: "Budget",
          displayName: "",
          approvalStatus: "",
          approvalDate: null,
          comment: "",
          userId: null,
          approvalString: "",
        },
        acctmgr1: {
          userLogin: "",
          jobTitle: "Accounting Manager 1",
          displayName: "",
          approvalStatus: "",
          approvalDate: new Date(),
          comment: "",
          userId: 0,
          approvalString: "",
        },
        acctmgr2: {
          userLogin: "",
          jobTitle: "Accounting Manager 2",
          displayName: "",
          approvalStatus: "",
          approvalDate: new Date(),
          comment: "",
          userId: 0,
          approvalString: "",
        },
        acctAdmin: {
          userLogin: "",
          jobTitle: "Accounting Admin",
          displayName: "",
          approvalStatus: "",
          approvalDate: new Date(),
          comment: "",
          userId: 0,
          approvalString: "",
        },

        //Everything below is only here for testing, Goal = Get rid of it
        // benefitToState: "",
        // costCenter: "",
        // domicile: "",
        // taNo: "",
        // departureTime: "",
        // returnTime: "",
        // fund: "",
        //
        // fYBudget: "0.00",
        // amtRemainBudget: "0.00",
        // amtRemainingAfterThis: "0.00",
        // authBudget: "0.00",
        //gL: "",
        //sMAGL: "",
        // fySpecialMarketing: "0.00",
        // fySpecialMarketingamtRemaining: "0.00",
        // fySpecialMarketingamtRemainingAfterThis: "0.00",
        // fYBudgetFY2: "0.00",
        // amtRemainBudgetFY2: "0.00",
        // amtRemainingAfterThisFY2: "0.00",
        // authBudgetFY2: "0.00",
        // fySpecialMarketingFY2: "0.00",
        // fySpecialMarketingamtRemainingFY2: "0.00",
        // fySpecialMarketingamtRemainingAfterThisFY2: "0.00",
        // airTravelAgencyUsed: null,
        // airTravelAgencyUsedJustification: "",
        // airFare: "",
        // vehicleType: "",
        // vehiclePassengers: "",
        // vehicleRentalTypeIsCompact: "",
        // vehicleRentalJustificationChoice: "",
        // vehicleRentalJustificationText: "",
        // limoTaxi: "",
        // limoTaxiFareAmount: "0.00",
        // tollsAndParking: "",
        // tollsAndParkingAmount: "0.00",
        // totalTransportationExpense: "0.00",
        // lodging: [
        //   {
        //     total: 0.0,
        //     days: 0.0,
        //     cost: 0.0,
        //   },
        //   {
        //     total: 0.0,
        //     days: 0.0,
        //     cost: 0.0,
        //   },
        //   {
        //     total: 0.0,
        //     days: 0.0,
        //     cost: 0.0,
        //   },
        // ],
        // meals: [
        //   {
        //     total: 0.0,
        //     days: 0.0,
        //     cost: 0.0,
        //   },
        // ],
        // tips: "",
        // tipsAmount: "0.00",
        // otherExpensePayableTo: "",
        // otherExpensePaymentMethod: "",
        // otherExpenseDueDate: "",
        // otherExpenseAmount: "0.00",
        // specialMarketingActivitiesAmountNotes: "",
        // travelAdvanceDate: "",
        // travelAdvanceAmount: "0.00",
        // chbxGPSRentalVehicle: false,
        // EstimatedCompensatoryTime: "",
        // budgetYear1: 0,
        // budgetYear2: 0,
      },

      hideDialog: true,
      dialogTitle: "",
      dialogText: "",
      requestID: "",
      formMode: "New",
      saving: false,
      printing: false,

      textInput: React.createRef(),
    };
    this.service = new DataService(this.props.context.pageContext);
    this.handleInput = this.handleInput.bind(this);
  }

  private handleInput(e) {
    let value = e.target.value;
    let name = e.target.name;
    this.setState((prevState) => ({
      reqData: {
        ...prevState.reqData,
        [name]: value,
      },
    }));
  }

  private handlereqDataTextChange(event) {
    const { name, value } = event.target;
    let reqData = { ...this.state.reqData };
    reqData[name] = value;
    this.setState({ reqData });
  }

  private _onFormatDate = (date: Date): string => {
    return date.toLocaleDateString();
  };

  private _onSelectDate(id, date: Date | null | undefined) {
    let reqData = { ...this.state.reqData };
    reqData[id] = date;
    this.setState({ reqData });
  }

  // private handleMaskedreqDataDateChange(event) {
  //   const { name, value } = event.target;
  //   let reqData = { ...this.state.reqData };
  //   let val = value.replace('_','');
  //   reqData[name] = value;
  //   reqData[name] = reqData[name].replace("_", "");
  //   this.setState({ reqData });
  // }

  private async handleMaskedDateWithValidation(event) {
    const { name, value } = event.target;
    let ctrlName = name;
    let reqData = { ...this.state.reqData };
    reqData[ctrlName] = value.replace(/_/g, "");
    let combinedDateTime = new Date();
    let valiMessage = "Required";
    let needToValidate = false;
    let testDate = new Date();
    // switch (ctrlName) {
    //   case "departureDateStr":
    //     combinedDateTime = new Date(
    //       reqData[ctrlName] + " " + reqData.departureTime
    //     );
    //     if (combinedDateTime.getTime() && reqData.departureTime) {
    //       reqData.departureDate = combinedDateTime;
    //       await this.setState({ DepartureDateError: "" });
    //     } else {
    //       valiMessage = "Departure Date and Time Required";
    //       await this.setState({
    //         DepartureDateError: "Valid Departure Date and Time Required",
    //       });
    //     }
    //     testDate = new Date(reqData[ctrlName]);
    //     if (!testDate.getTime() || !reqData[ctrlName]) {
    //       needToValidate = true;
    //     }
    //     break;

    //   case "departureTime":
    //     combinedDateTime = new Date(
    //       reqData.departureDateStr + " " + reqData[ctrlName]
    //     );
    //     if (combinedDateTime.getTime() && reqData.departureTime) {
    //       reqData.departureDate = combinedDateTime;
    //       await this.setState({ DepartureDateError: "" });
    //     } else {
    //       valiMessage = "Valid Departure Date and Time Required";
    //       await this.setState({
    //         DepartureDateError: "Valid Departure Date and Time Required",
    //       });
    //     }
    //     testDate = new Date("9/9/2009 " + reqData[ctrlName]);
    //     if (!testDate.getTime() || !reqData[ctrlName]) {
    //       needToValidate = true;
    //     }
    //     break;

    //   case "returnDateStr":
    //     combinedDateTime = new Date(
    //       reqData[ctrlName] + " " + reqData.returnTime
    //     );
    //     if (combinedDateTime.getTime() && reqData.returnTime) {
    //       reqData.returnDate = combinedDateTime;
    //       await this.setState({ ReturnDateError: "" });
    //     } else {
    //       valiMessage = "Valid Return Date and Time Required";
    //       await this.setState({
    //         ReturnDateError: "Valid Return Date and Time Required",
    //       });
    //     }
    //     testDate = new Date(reqData[ctrlName]);
    //     if (!testDate.getTime() || !reqData[ctrlName]) {
    //       needToValidate = true;
    //     }
    //     break;

    //   case "returnTime":
    //     combinedDateTime = new Date(
    //       reqData.returnDateStr + " " + reqData[ctrlName]
    //     );
    //     if (combinedDateTime.getTime() && reqData.returnTime) {
    //       reqData.returnDate = combinedDateTime;
    //       await this.setState({ ReturnDateError: "" });
    //     } else {
    //       valiMessage = "Valid Return Date and Time Required";
    //       await this.setState({
    //         ReturnDateError: "Valid Return Date and Time Required",
    //       });
    //     }
    //     testDate = new Date("9/9/2009 " + reqData[ctrlName]);
    //     if (!testDate.getTime() || !reqData[ctrlName]) {
    //       needToValidate = true;
    //     }
    //     break;
    // }
    await this.setState({ reqData });

    await this.setState((prevState) => {
      const validations = prevState.validations.filter(
        (vali) => vali.controlName !== ctrlName
      );
      if (needToValidate) {
        validations.push({
          controlName: ctrlName,
          message: valiMessage,
        });
      }
      return { validations };
    });
  }

  //Used for general Checkboxes
  private _onControlledCheckboxChange(event) {
    const { name, checked } = event.target;
    let reqData = { ...this.state.reqData };
    reqData[name] = checked;
    this.setState({ reqData });
  }

  //Used for yes/No checkboxes, Unchecks other box
  private async _onUniqueCheckboxChange(checkboxVal, event) {
    const { name } = event.target;
    let reqData = { ...this.state.reqData };
    checkboxVal = reqData[name] == checkboxVal ? "" : checkboxVal; //to allow for unchecking
    reqData[name] = checkboxVal;

    //If No car rental, change cost to 0
    if (reqData.carRentalUsed == "false") {
      reqData.vehicleRentalCost = "0.00";
    }

    await this.setState({ reqData });
    this.updateCurrencyCalculations(this);
  }

  //Used to handle all data changes regarding numbers (Mainly section C)
  private async handlereqDataNumberChange(fieldName, value) {
    let reqData = { ...this.state.reqData };
    //let val = !isNaN(value.floatValue) ? value.floatValue : "";
    let val = Number(value.target.value.replace(/[^0-9\.]+/g, ""))
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");
    //let val = !isNaN(temp) ? parseFloat(temp) : "";
    reqData[fieldName] = val;
    await this.setState({ reqData });
    this.updateCurrencyCalculations(this);
  }
  //Used to handle all data changes regarding numbers (Mainly section C)
  private async handlereqDataWholeNumberChange(fieldName, value) {
    let reqData = { ...this.state.reqData };
    //let val = !isNaN(value.floatValue) ? value.floatValue : "";
    let val = Number(value.target.value.replace(/[^0-9\.]+/g, "")).toFixed();
    //let val = !isNaN(temp) ? parseFloat(temp) : "";
    reqData[fieldName] = val;
    await this.setState({ reqData });
    this.updateCurrencyCalculations(this);
  }

  private printPage() {
    window.print();
  }

  private updateCurrencyCalculations(ctx) {
    let reqData = { ...this.state.reqData };

    //trim trailing decimals
    let miles = reqData.mileageEstimation ? reqData.mileageEstimation : 0.0;
    let airFare = reqData.airFareCost
      ? Number(reqData.airFareCost.replace(/,/g, ""))
      : 0.0;
    let vehicleRentalCost = reqData.vehicleRentalCost
      ? Number(reqData.vehicleRentalCost.replace(/,/g, ""))
      : 0.0;
    // let limoTaxiFareAmount = reqData.limoTaxiFareAmount
    //   ? Number(reqData.limoTaxiFareAmount.replace(/,/g, ""))
    //   : 0.0;
    let mileageRate = reqData.mileageRate ? Number(reqData.mileageRate) : 0.0;

    //Updating Mileage Amount
    reqData.mileageAmount = (
      Number(reqData.mileageEstimation.replace(/,/g, "")) * reqData.mileageRate
    )
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");

    //Updating Total Lodging Costs
    reqData.totalLodgingAmount = (
      Number(reqData.lodgingCostPerNight.replace(/,/g, "")) *
      Number(reqData.lodgingNights.replace(/,/g, ""))
    )
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");

    //Updating Total Lodging Costs
    reqData.totalMealAmount = (
      Number(reqData.mealCostPerNight.replace(/,/g, "")) *
      Number(reqData.mealPerNights.replace(/,/g, ""))
    )
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");

    //Updating Cost Per Traveler
    reqData.costPerTraveler = (
      Number(reqData.registrationFees.replace(/,/g, "")) +
      Number(reqData.airFareCost.replace(/,/g, "")) +
      Number(reqData.mileageAmount.replace(/,/g, "")) +
      Number(reqData.totalLodgingAmount.replace(/,/g, "")) +
      Number(reqData.totalMealAmount.replace(/,/g, "")) +
      Number(reqData.vehicleRentalCost.replace(/,/g, "")) +
      Number(reqData.otherTransportCosts.replace(/,/g, ""))
    )
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");

    //Configuring Special marketing activity to make it compatible
    let specialMarketingActivitiesAmount =
      reqData.specialMarketingActivitiesAmount
        ? Number(reqData.specialMarketingActivitiesAmount.replace(/,/g, ""))
        : 0.0;
    //Total Estimated cost of trip
    reqData.totalEstimatedCostOfTrip = (
      Number(reqData.costPerTraveler.replace(/,/g, "")) *
        Number(reqData.numberOfTravelers.replace(/,/g, "")) +
      specialMarketingActivitiesAmount
    )
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");

    //set state
    this.setState({ reqData });
  }

  private genericValidation(
    ctrlName: string,
    isNotValid: boolean,
    message: string,
    value: string
  ) {
    //uses:
    //if message left blank and  invalid, message will be created
    //example: onGetErrorMessage={this.genericValidation.bind(this,name,this.state.customProp!=='sparkhound')}
    let valiMessage = message ? message : "Invalid";

    //if isNotValid condition is set to null, require
    //example: onGetErrorMessage={this.genericValidation.bind(this,name,null,'Need To Validate)}
    let needToValidate = isNotValid ? isNotValid : stringIsNullOrEmpty(value);

    this.setState((prevState) => {
      const validations = prevState.validations.filter(
        (vali) => vali.controlName !== ctrlName
      );
      if (needToValidate) {
        validations.push({
          controlName: ctrlName,
          message: valiMessage,
        });
      }
      return { validations };
    });

    //return message to set control validation
    return needToValidate ? valiMessage : "";
  }

  private async _getPeoplePickerItems(items: any[]) {
    if (items.length > 0) {
      let selectedUser = await sp.web.ensureUser(items[0].id);
      await this.setState((prevState) => ({
        reqData: {
          ...prevState.reqData,
          employeeId: selectedUser.data.Id,
          employeeName: items[0].text,
          employeeLogin: items[0].id,
        },
      }));
      this._getAndSetApprovers();
    } else {
      this.setState((prevState) => ({
        reqData: {
          ...prevState.reqData,
          employeeId: undefined,
          employeeName: "",
          employeeLogin: "",
        },
      }));
    }
  }

  private async _getAndSetApprovers() {
    let employeesApprovers = await this.service.GetApprovers(
      this.state.reqData.employeeLogin
    );
    console.log(employeesApprovers);
    let reqData = { ...this.state.reqData };

    reqData.employeeApproval.displayName = reqData.employeeName
      ? reqData.employeeName
      : "";
    reqData.employeeApproval.userLogin = reqData.employeeLogin
      ? reqData.employeeLogin.split("|")[2].toLowerCase()
      : "";
    reqData.employeeApproval.userId = reqData.employeeId
      ? reqData.employeeId
      : null;

    if (reqData.employeeApproval.approvalStatus == "Approved") {
      reqData.employeeApproval.approvalString = `Approved by ${
        reqData.employeeApproval.displayName
      } at ${new Date(
        reqData.employeeApproval.approvalDate.toString()
      ).toDateString()} ${new Date(
        reqData.employeeApproval.approvalDate.toString()
      ).toLocaleTimeString()}`;
    } else if (reqData.employeeApproval.approvalStatus == "") {
      reqData.sectionHead.approvalString = `Pending Approval from ${reqData.employeeLogin
        .split("|")[2]
        .toLowerCase()}`;
    }

    reqData.sectionHead.displayName = employeesApprovers.SectionHead
      ? employeesApprovers.SectionHead.Title
      : "";
    reqData.sectionHead.userLogin = employeesApprovers.SectionHead
      ? employeesApprovers.SectionHead.UserName.toLowerCase()
      : "";
    reqData.sectionHead.userId = employeesApprovers.SectionHead
      ? employeesApprovers.SectionHead.Id
      : null;
    if (!employeesApprovers.SectionHead) {
      reqData.sectionHead.approvalString = "N/A";
    } else if (reqData.sectionHead.approvalStatus == "Approved") {
      reqData.sectionHead.approvalString = `Approved by ${
        employeesApprovers.SectionHead.Title
      } at ${new Date(
        reqData.sectionHead.approvalDate.toString()
      ).toDateString()} ${new Date(
        reqData.sectionHead.approvalDate.toString()
      ).toLocaleTimeString()}`;
    } else if (reqData.sectionHead.approvalStatus == "") {
      reqData.sectionHead.approvalString = `Pending Approval from ${employeesApprovers.SectionHead.Title}`;
    }

    reqData.secretary.displayName = employeesApprovers.Secretary
      ? employeesApprovers.Secretary.Title
      : "";
    reqData.secretary.userLogin = employeesApprovers.Secretary
      ? employeesApprovers.Secretary.UserName.toLowerCase()
      : "";
    reqData.secretary.userId = employeesApprovers.Secretary
      ? employeesApprovers.Secretary.Id
      : null;
    if (!employeesApprovers.Secretary) {
      reqData.secretary.approvalString = "N/A";
    } else if (reqData.secretary.approvalStatus == "Approved") {
      reqData.secretary.approvalString = `Approved by ${
        employeesApprovers.Secretary.Title
      } at ${new Date(
        reqData.secretary.approvalDate.toString()
      ).toDateString()} ${new Date(
        reqData.secretary.approvalDate.toString()
      ).toLocaleTimeString()}`;
    } else if (reqData.secretary.approvalStatus == "") {
      reqData.secretary.approvalString = `Pending Approval from ${employeesApprovers.Secretary.Title}`;
    }

    reqData.undersecretary.displayName = employeesApprovers.Undersecretary
      ? employeesApprovers.Undersecretary.Title
      : "";
    reqData.undersecretary.userLogin = employeesApprovers.Undersecretary
      ? employeesApprovers.Undersecretary.UserName.toLowerCase()
      : "";
    reqData.undersecretary.userId = employeesApprovers.Undersecretary
      ? employeesApprovers.Undersecretary.Id
      : null;
    if (!employeesApprovers.Undersecretary) {
      reqData.undersecretary.approvalString = "N/A";
    } else if (reqData.undersecretary.approvalStatus == "Approved") {
      reqData.undersecretary.approvalString = `Approved by ${
        employeesApprovers.Undersecretary.Title
      } at ${new Date(
        reqData.undersecretary.approvalDate.toString()
      ).toDateString()} ${new Date(
        reqData.undersecretary.approvalDate.toString()
      ).toLocaleTimeString()}`;
    } else if (reqData.undersecretary.approvalStatus == "") {
      reqData.undersecretary.approvalString = `Pending Approval from ${employeesApprovers.Undersecretary.Title}`;
    }

    reqData.deputyUndersecretary.displayName =
      employeesApprovers.DeputyUndersecretary
        ? employeesApprovers.DeputyUndersecretary.Title
        : "";
    reqData.deputyUndersecretary.userLogin =
      employeesApprovers.DeputyUndersecretary
        ? employeesApprovers.DeputyUndersecretary.UserName.toLowerCase()
        : "";
    reqData.deputyUndersecretary.userId =
      employeesApprovers.DeputyUndersecretary
        ? employeesApprovers.DeputyUndersecretary.Id
        : null;
    if (!employeesApprovers.DeputyUndersecretary) {
      reqData.deputyUndersecretary.approvalString = "N/A";
    } else if (reqData.deputyUndersecretary.approvalStatus == "Approved") {
      reqData.deputyUndersecretary.approvalString = `Approved by ${
        employeesApprovers.DeputyUndersecretary.Title
      } at ${new Date(
        reqData.deputyUndersecretary.approvalDate.toString()
      ).toDateString()} ${new Date(
        reqData.deputyUndersecretary.approvalDate.toString()
      ).toLocaleTimeString()}`;
    } else if (reqData.deputyUndersecretary.approvalStatus == "") {
      reqData.deputyUndersecretary.approvalString = `Pending Approval from ${employeesApprovers.DeputyUndersecretary.Title}`;
    }

    reqData.budget.displayName = employeesApprovers.Budget
      ? employeesApprovers.Budget.Title
      : "";
    reqData.budget.userLogin = employeesApprovers.Budget
      ? employeesApprovers.Budget.UserName.toLowerCase()
      : "";
    reqData.budget.userId = employeesApprovers.Budget
      ? employeesApprovers.Budget.Id
      : null;

    reqData.acctmgr1.displayName = employeesApprovers.AcctMgr1
      ? employeesApprovers.AcctMgr1.Title
      : "";
    reqData.acctmgr1.userLogin = employeesApprovers.AcctMgr1
      ? employeesApprovers.AcctMgr1.UserName.toLowerCase()
      : "";

    reqData.acctmgr2.displayName = employeesApprovers.AcctMgr2
      ? employeesApprovers.AcctMgr2.Title
      : "";
    reqData.acctmgr2.userLogin = employeesApprovers.AcctMgr2
      ? employeesApprovers.AcctMgr2.UserName.toLowerCase()
      : "";

    //reqData.acctAdmin.userLogin = employeesApprovers.AcctAdmin ? employeesApprovers.AcctAdmin.UserName.toLowerCase() : "";

    reqData.agency = employeesApprovers.Agency ? employeesApprovers.Agency : "";
    reqData.personnelNo = employeesApprovers.PersonnelNo
      ? employeesApprovers.PersonnelNo
      : "";
    this.setState({ reqData });
  }
  //When someone click button, fires off
  private async approvalButton(approvalName) {
    //set approval status on the current approval object

    // if(st.reqData.employeeApproval.approvalString = ""){
    //   st.reqData.employeeApproval.approvalStatus = "Approved";
    //   st.reqData.employeeApproval.approvalDate = new Date();
    //   st.reqData.employeeApproval.approvalString = `Approved by ${st.reqData.employeeApproval.displayName} at ${new Date().toDateString()} ${new Date().toLocaleTimeString()}`
    //   await this.setState(st);
    // }
    var st = { ...this.state };
    var skipApprovalVerbiage = "N/A";
    st.reqData.status = "In Progress";
    st.reqData[approvalName].approvalStatus = "Approved";
    st.reqData[approvalName].approvalDate = new Date();
    st.reqData[approvalName].approvalString = `Approved by ${
      st.reqData[approvalName].displayName
    } at ${new Date().toDateString()} ${new Date().toLocaleTimeString()}`;
    await this.setState(st);
    //set request stage based on the next approver  and  set next approver field (person)

    switch (approvalName) {
      case "employeeApproval":
        if (
          st.reqData.sectionHead.approvalStatus == "" &&
          st.reqData.sectionHead.userId
        ) {
          st.reqData.stage = "Section Head";
          st.reqData.nextApprover = st.reqData.sectionHead.userId;
        } else if (
          st.reqData.secretary.approvalStatus == "" &&
          st.reqData.secretary.userId
        ) {
          st.reqData.stage = "Department Head";
          st.reqData.nextApprover = st.reqData.secretary.userId;
          st.reqData.sectionHead.approvalStatus = skipApprovalVerbiage;
          st.reqData.sectionHead.approvalString = skipApprovalVerbiage;
          st.reqData.sectionHead.approvalDate = new Date();
        } else if (
          st.reqData.undersecretary.approvalStatus == "" &&
          st.reqData.undersecretary.userId
        ) {
          st.reqData.stage = "Undersecretary";
          st.reqData.nextApprover = st.reqData.undersecretary.userId;
          st.reqData.sectionHead.approvalStatus = skipApprovalVerbiage;
          st.reqData.sectionHead.approvalString = skipApprovalVerbiage;
          st.reqData.sectionHead.approvalDate = new Date();
          st.reqData.secretary.approvalStatus = skipApprovalVerbiage;
          st.reqData.secretary.approvalString = skipApprovalVerbiage;
          st.reqData.secretary.approvalDate = new Date();
        }
        break;

      case "sectionHead":
        if (
          st.reqData.secretary.approvalStatus == "" &&
          st.reqData.secretary.userId
        ) {
          st.reqData.stage = "Department Head";
          st.reqData.nextApprover = st.reqData.secretary.userId;
        } else if (
          st.reqData.undersecretary.approvalStatus == "" &&
          st.reqData.secretary.userLogin == "" &&
          st.reqData.undersecretary.userId
        ) {
          st.reqData.stage = "Undersecretary";
          st.reqData.nextApprover = st.reqData.undersecretary.userId;
          st.reqData.secretary.approvalStatus = skipApprovalVerbiage;
          st.reqData.secretary.approvalString = skipApprovalVerbiage;
          st.reqData.secretary.approvalDate = new Date();
        } else if (
          st.reqData.budget.approvalStatus == "" &&
          st.reqData.undersecretary.userLogin == "" &&
          st.reqData.budget.userId
        ) {
          st.reqData.stage = "Budget";
          st.reqData.nextApprover = st.reqData.budget.userId;
          st.reqData.secretary.approvalStatus = skipApprovalVerbiage;
          st.reqData.secretary.approvalString = skipApprovalVerbiage;
          st.reqData.secretary.approvalDate = new Date();

          st.reqData.undersecretary.approvalStatus = skipApprovalVerbiage;
          st.reqData.undersecretary.approvalString = skipApprovalVerbiage;
          st.reqData.undersecretary.approvalDate = new Date();
        }
        break;

      case "secretary":
        //Secretary
        if (
          st.reqData.undersecretary.approvalStatus == "" &&
          st.reqData.undersecretary.userId
        ) {
          st.reqData.stage = "Undersecretary";
          st.reqData.nextApprover = st.reqData.undersecretary.userId;
        } else if (
          st.reqData.undersecretary.userLogin == "" &&
          st.reqData.budget.approvalStatus == ""
        ) {
          st.reqData.stage = "Budget";
          st.reqData.nextApprover = st.reqData.budget.userId;
          st.reqData.undersecretary.approvalStatus = skipApprovalVerbiage;
          st.reqData.undersecretary.approvalString = skipApprovalVerbiage;
          st.reqData.undersecretary.approvalDate = new Date();
        }
        break;

      case "undersecretary":
        //Undersecretary
        if (
          st.reqData.budget.approvalStatus == "" &&
          st.reqData.budget.userId
        ) {
          st.reqData.stage = "Budget";
          st.reqData.nextApprover = st.reqData.budget.userId;
        } else if (
          st.reqData.budget.userLogin == "" &&
          st.reqData.deputyUndersecretary.approvalStatus == "" &&
          st.reqData.deputyUndersecretary.userId
        ) {
          st.reqData.stage = "Deputy Undersecretary";
          st.reqData.nextApprover = st.reqData.deputyUndersecretary.userId;
          st.reqData.budget.approvalStatus = skipApprovalVerbiage;
          st.reqData.budget.approvalString = skipApprovalVerbiage;
          st.reqData.budget.approvalDate = new Date();
        }
        break;

      case "budget":
        //Budget
        if (
          st.reqData.deputyUndersecretary.approvalStatus == "" &&
          st.reqData.deputyUndersecretary.userId
        ) {
          st.reqData.stage = "Deputy Undersecretary";
          st.reqData.nextApprover = st.reqData.deputyUndersecretary.userId;
        } else if (st.reqData.deputyUndersecretary.userLogin == "") {
          st.reqData.deputyUndersecretary.approvalStatus = skipApprovalVerbiage;
          st.reqData.deputyUndersecretary.approvalString = skipApprovalVerbiage;
          st.reqData.undersecretary.approvalDate = new Date();

          st.reqData.stage = "Complete";
          st.reqData.nextApprover = null;
        } else if (
          st.reqData.deputyUndersecretary.approvalStatus == "Approved"
        ) {
          st.reqData.stage = "Complete";
          st.reqData.nextApprover = null;
        }
        break;

      case "deputyUndersecretary":
        //Deputy Undersecretary
        st.reqData.stage = "Complete";
        st.reqData.nextApprover = null;
        break;
    }

    await this.setState({ kickoffFLOW: "Yes" });

    //append approval info to request log
    st.reqData.requestLog = `${st.reqData.requestLog} \n${
      st.reqData[approvalName].displayName
    } (login: ${st.reqData[approvalName].userLogin}) approved at ${st.reqData[
      approvalName
    ].approvalDate.toDateString()} ${st.reqData[
      approvalName
    ].approvalDate.toLocaleTimeString()}`;

    //prompt to save the form or continue
    //st.dialogTitle = "Approval";  //removing prompt per LED request
    //st.dialogText = "Do you wish to save and close your approval or stay on the page?";
    //st.hideDialog = false;
    await this.setState(st);
    await this.SaveAndCloseButton("Yes");

    //Lets user know if unapproved specs are not approved
    // let specApprArray = [
    //   "chbxVehicleRental",
    //   "chbxGPSRentalVehicle",
    //   "chbxProspectInSameHotelAsEmployee",
    //   "chbxSpecialMarketingActivities",
    //   "chbx50pctLodgingException",
    //   "chbxOther",
    // ];
    // let unapprovedSpecs = false;
    // specApprArray.forEach((e) => {
    //   if (
    //     st.reqData[e] == true &&
    //     st.reqData[e + "Sig"] == "" &&
    //     (approvalName == "deputyUndersecretary" ||
    //       approvalName == "undersecretary")
    //   ) {
    //     unapprovedSpecs = true;
    //   }
    // });
    // if (unapprovedSpecs) {
    //   //toast.success("One or more special approvals still need approval! Please Approve and Save Form");
    //   toast.warn(
    //     "One or items in the Special Approvals Section still need approval! Please Approve and Save Form"
    //   );
    // } else {
    //   toast.success("Form approved!");
    // }
  }

  //private rejectButton(approval: Approver, event) {
  //check and only continue if comment is added

  //prompt to ensure that user wants to cancel existing approvals and restart process

  //remove approval status on all other approvals and set stage

  //set rejection info with user name, date and comments to be used in emails and logs

  //append rejection info to request log

  //save form
  //this.SaveButton();
  //}

  private async SaveButton() {
    this.setState({ saving: true });
    //toast.success("Saving Item to Service Database!");
    let itemId = await this.service.SaveRequest(this.state);
    //toast.success("Form not saved yet 2!");
    this.setState({ saving: false, requestID: itemId });
    toast.success("Form saved!");
  }

  private async emailPDF() {
    this.setState({ saving: true });
    let itemId = await this.service.SaveRequest(this.state);
    let itemEmailReqId = await this.service.SaveEmailSubmission(
      this.state.requestID
    );
    this.setState({ saving: false, requestID: itemId });
    toast.success("Form email request submitted!");
  }

  private async SaveAndCloseButton(kFV?) {
    this.setState({ saving: true });
    let itemId = await this.service.SaveRequest(this.state, kFV);
    this.setState({ saving: false, requestID: itemId });
    //this.CloseForm();
  }

  private async Submit() {
    this.approvalButton("employeeApproval");

    this.setState({ saving: true });
    let itemId = await this.service.SaveRequest(this.state);
    this.setState({ saving: false, requestID: itemId });

    this.CloseForm();
    // if (this.state.reqData.status == "Draft") {
    //   let reqData = { ...this.state.reqData };
    //   reqData.status = "In Progress";
    //   let kickoffFlowValue = "Yes";
    //   this.setState({ kickoffFLOW: "Yes" });

    //   if (
    //     reqData.employeeApproval.userLogin ==
    //       this.props.context.pageContext.user.loginName.toLowerCase() &&
    //     reqData.employeeApproval.approvalStatus == ""
    //   ) {
    //     this.approvalButton("employeeApproval");
    //   } else if (reqData.employeeApproval.approvalStatus == "") {
    //     reqData.stage = "Employee Approval";
    //     reqData.nextApprover = reqData.employeeApproval.userId;
    //     await this.setState({ reqData });
    //     await this.service.SaveRequest(this.state, kickoffFlowValue);
    //   }
    // } else {
    //   await this.service.SaveRequest(this.state);
    // }
  }

  private CloseForm() {
    let queryParameters = new UrlQueryParameterCollection(window.location.href);
    const srcUrl = queryParameters.getValue("Source");
    let webUrl = this.props.context.pageContext.web.absoluteUrl;
    window.location.href = srcUrl ? srcUrl : webUrl;
  }

  //handle adding attachments
  // private showForm() {
  //   //let reqData = { ...this.state.reqData };
  //   //reqData.Adding= true;
  //   //this.setState({ reqData });
  //   this.setState({ AddingAttachment: true });
  // }

  private _onClose(success) {
    this.setState({ AddingAttachment: false });
    setTimeout(
      function () {
        this.FetchAttachments();
      }.bind(this),
      1000
    );
  }

  private async FetchAttachments() {
    let results = await this.service.GetAttachments(this.state.reqData.formKey);
    //let libPath = this.props.context.web
    this.setState({ Attachments: results });
  }

  // private async RemoveAttachment(attachment) {
  //   await this.service.RemoveAttachment(attachment.Id);
  //   setTimeout(
  //     function () {
  //       this.FetchAttachments();
  //     }.bind(this),
  //     1000
  //   );
  // }

  public componentDidMount() {
    this.init();
  }
  private async init() {
    //is Edit or New
    let queryParameters = new UrlQueryParameterCollection(window.location.href);
    const requestID = queryParameters.getValue("RequestID");
    if (requestID) {
      let tempState = { ...this.state };
      tempState.requestID = requestID;
      tempState.formMode = "Edit";
      this.setState(tempState);
      //get old data from item to populate the reqData object
      try {
        let results = await this.service.getRequestData(tempState.requestID);
        let parsedReqData = JSON.parse(results["RequestData"]);
        //parsedReqData.otherExpenseDueDate = parsedReqData.otherExpenseDueDate ? new Date(parsedReqData.otherExpenseDueDate):null;
        //parsedReqData.departureDate = parsedReqData.departureDate ? new Date(parsedReqData.departureDate): null;
        parsedReqData.dateOfRequest = parsedReqData.dateOfRequest
          ? new Date(parsedReqData.dateOfRequest)
          : null;
        //parsedReqData.returnDate = parsedReqData.returnDate ? new Date(parsedReqData.returnDate): null;
        //parsedReqData.travelAdvanceDate = parsedReqData.travelAdvanceDate ? new Date(parsedReqData.travelAdvanceDate): null;
        parsedReqData.taNo = requestID ? requestID : "";
        parsedReqData.budgetYear1 = parsedReqData.budgetYear1
          ? parsedReqData.budgetYear1
          : Number(this.props.startingFinancialYear.toString().slice(-2));
        parsedReqData.budgetYear2 = parsedReqData.budgetYear2
          ? parsedReqData.budgetYear2
          : Number(this.props.startingFinancialYear.toString().slice(-2)) + 1;
        this.setState({ reqData: parsedReqData });
        this.FetchAttachments();
        //this.setState({ results }, () => {
        //});
      } catch (error) {}
      //get approvers
    } else {
      //set defaults for new form
      let curUserId = await sp.web.ensureUser(
        this.props.context.pageContext.user.loginName
      );
      let data = { ...this.state.reqData };
      data.employeeId = curUserId.data.Id;
      data.employeeName = curUserId.data.Title;
      data.employeeLogin = curUserId.data.LoginName;
      data.formKey = getRandomString(8);
      data.mileageRate = this.props.mileageRate
        ? Number(this.props.mileageRate)
        : 0.575;
      // data.budgetYear1 = data.budgetYear1
      //   ? data.budgetYear1
      //   : Number(this.props.startingFinancialYear.toString().slice(-2));
      // data.budgetYear2 = data.budgetYear2
      //   ? data.budgetYear2
      //   : Number(this.props.startingFinancialYear.toString().slice(-2)) + 1;
      this.setState({ reqData: data });
      //this._addMultiDay('meals', null);
      //this._addMultiDay('lodging', null);
    }

    this._getAndSetApprovers();

    let reqData = { ...this.state.reqData };

    // reqData.employeeApproval.userLogin = "admin@laecondev.onmicrosoft.com";
    // reqData.sectionHead.userLogin = "admin@laecondev.onmicrosoft.com";
    // reqData.secretary.userLogin = "admin@laecondev.onmicrosoft.com";
    // reqData.undersecretary.userLogin = "admin@laecondev.onmicrosoft.com";
    // reqData.deputyUndersecretary.userLogin = "admin@laecondev.onmicrosoft.com";
    // reqData.budget.userLogin = "admin@laecondev.onmicrosoft.com";
    // reqData.acctmgr1.userLogin = "admin@laecondev.onmicrosoft.com";
    // reqData.acctmgr2.userLogin = "admin@laecondev.onmicrosoft.com";
    // reqData.acctAdmin.userLogin = "admin@laecondev.onmicrosoft.com";

    this.setState({ reqData });

    //end Init
  }
  public render(): React.ReactElement<ITravelRequestProps> {
    const { error, message, results, reqData, validations, AddingAttachment } =
      this.state;
    const {
      sectionHead,
      secretary,
      undersecretary,
      deputyUndersecretary,
      budget,
      employeeLogin,
      acctmgr1,
      acctmgr2,
      acctAdmin,
    } = this.state.reqData;
    const currentUser = this.props.context.pageContext.user;
    const addIcon: IIconProps = { iconName: "Add" };
    const removeIcon: IIconProps = { iconName: "Cancel" };
    let disableSubmit = validations.length > 0 ? true : false;
    let disableSubmitForSpecialSigs = false; //Change to false if you want to use this function
    if (
      !reqData.agency ||
      !reqData.departureDateStr ||
      !reqData.returnDateStr ||
      !reqData.destination ||
      !reqData.justficationForTrip
      //!reqData.personnelNo ||
      // !reqData.benefitToState
      // !reqData.domicile
    ) {
      disableSubmit = true;
    }
    // if (
    //   (reqData.stage == "Secretary" ||
    //     reqData.stage == "Undersecretary" ||
    //     reqData.stage == "Deputy Undersecretary") &&
    //   ((reqData.chbx50pctLodgingException &&
    //     !reqData.chbx50pctLodgingExceptionSig) ||
    //     (reqData.chbxOther && !reqData.chbxOtherSig) ||
    //     (reqData.chbxProspectInSameHotelAsEmployee &&
    //       !reqData.chbxProspectInSameHotelAsEmployeeSig) ||
    //     (reqData.chbxSpecialMarketingActivities &&
    //       !reqData.chbxSpecialMarketingActivitiesSig))
    // ) {
    //   disableSubmitForSpecialSigs = true;
    // }

    const isBudgetApprover =
      budget.userLogin == currentUser.loginName.toLowerCase() ? true : false;
    //Controls Section F fields
    const isApprover =
      sectionHead.userLogin == currentUser.loginName.toLowerCase() ||
      secretary.userLogin.toLowerCase() == currentUser.loginName ||
      undersecretary.userLogin.toLowerCase() == currentUser.loginName ||
      deputyUndersecretary.userLogin.toLowerCase() == currentUser.loginName
        ? true
        : false;
    const isAcctMgr =
      acctmgr1.userLogin.toLowerCase() == currentUser.loginName ||
      acctmgr2.userLogin.toLowerCase() == currentUser.loginName
        ? true
        : false;
    //const isAdmin = acctAdmin.userLogin == currentUser.loginName ? true : false;
    const isAdmin =
      "molly.hendricks@laecondev.onmicrosoft.com" == currentUser.loginName ||
      "kristin.pace@laecondev.onmicrosoft.com" == currentUser.loginName ||
      "nicolaus.james@laecondev.onmicrosoft.com" == currentUser.loginName ||
      "nicolaus.james@la.gov" == currentUser.loginName ||
      "admin@laecondev.onmicrosoft.com" == currentUser.loginName
        ? true
        : false;
    const disableControls =
      reqData.status == "Draft" || isAcctMgr || isAdmin ? false : true;
    //const empMinusClaims = employeeLogin ? employeeLogin.split('|')[2] : currentUser.loginName;
    const empMinusClaims = employeeLogin ? [employeeLogin.split("|")[2]] : [];

    return (
      <div className={`${styles.travelRequest} printarea`}>
        <ToastContainer position="bottom-center" hideProgressBar={true} />
        <div className="form-group">
          {/* Header Text */}
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <h2 className="col align-self-center">Travel Request Form</h2>
            </div>
            <div className="ms-Grid-row">
              <p>
                NO REGISTRATIONS OR RESERVATIONS SHOULD BE MADE UNTIL ALL
                APPROVALS ARE OBTAINED Instructions: Complete all sections
                pertaining to your request. Print or Type all entries. Submit
                completed form with all necessary approvals to your Agency???s
                Travel Administrator. Retain a copy for your records.
              </p>
            </div>
          </div>

          {/* Section A*/}
          <div className="ms-Grid">
            {/* Section A Row 1*/}
            <div className="ms-Grid-row">
              <h2 className={styles.sectionHeader}>
                Section A: General Information- Complete All Info
              </h2>
            </div>

            {/* Section A Row 2*/}
            <Stack horizontal>
              <label className={styles.generalLabel}>Name:</label>
              <div className="ms-Grid-col ms-sm4">
                <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={1}
                  peoplePickerCntrlclassName="slimPeoplePicker"
                  showtooltip={true}
                  isRequired={true}
                  defaultSelectedUsers={empMinusClaims}
                  disabled={disableControls}
                  selectedItems={this._getPeoplePickerItems.bind(this)}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={400}
                />
              </div>
              {/* <TextField
                underlined
                className="ms-Grid-col ms-sm6"
                name="employeeName"
                value={reqData.employeeName}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty(reqData.employeeName),
                  "Name Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              /> */}
              <div className="ms-Grid-col ms-sm4">
                <DatePicker
                  className={styles.DatepickerComboBox}
                  label="Date of Request: "
                  id="dateOfRequest"
                  isRequired={true}
                  strings={DatePickerStrings}
                  //dateConvention={DateConvention.Date}
                  underlined
                  formatDate={this._onFormatDate}
                  value={reqData.dateOfRequest}
                  //disabled={ disableControls }
                  disabled={true}
                  onSelectDate={this._onSelectDate.bind(this, "dateOfRequest")}
                />
              </div>
              &nbsp;
              <label className={styles.generalLabel}>Destination:</label>
              <TextField
                underlined
                className="ms-Grid-col ms-sm4"
                name="destination"
                value={reqData.destination}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty(reqData.destination),
                  "Destination Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              />
            </Stack>

            {/* Section A Row 3*/}
            <br></br>
            <Stack horizontal>
              <label className={styles.generalLabel}>Title:</label>
              <TextField
                underlined
                className="ms-Grid-col ms-sm8"
                name="employeeTitle"
                value={reqData.employeeTitle}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty(reqData.employeeTitle),
                  "Title Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              />
              &nbsp;
              <label className={styles.generalLabel}>Begin Date:</label>
              <MaskedInput
                mask="11/11/1111"
                name="departureDateStr"
                onChange={this.handleMaskedDateWithValidation.bind(this)}
                value={reqData.departureDateStr}
                className={styles.inputMaskControl}
                disabled={disableControls}
                required={true}
              />
              <label className={styles.generalLabel}>End Date:</label>
              <MaskedInput
                mask="11/11/1111"
                name="returnDateStr"
                onChange={this.handleMaskedDateWithValidation.bind(this)}
                value={reqData.returnDateStr}
                className={styles.inputMaskControl}
                disabled={disableControls}
                required={true}
              />
            </Stack>

            {/* Section A Row 4*/}
            <br></br>
            <Stack horizontal>
              <label className={styles.generalLabel}>Agency:</label>
              <TextField
                underlined
                className="ms-Grid-col ms-sm2"
                name="agency"
                value={reqData.agency}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty(reqData.agency),
                  "Agency Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              />
              &nbsp;
              <label className={styles.generalLabel}>Division/Section:</label>
              <TextField
                underlined
                className="ms-Grid-col ms-sm3"
                name="division"
                value={reqData.division}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty(reqData.division),
                  "Division/Section Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              />
              &nbsp;
              <label className={styles.generalLabel}>
                Mode of Transportation:
              </label>
              <TextField
                underlined
                className="ms-Grid-col ms-sm3"
                name="modeOfTransportation"
                value={reqData.modeOfTransportation}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty(reqData.modeOfTransportation),
                  "Mode Of Transportation Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              />
            </Stack>

            {/* Section A Row 5*/}
            <br></br>
            <Stack>
              <label className={styles.generalLabel}>
                Justification for trip:
              </label>
              <TextField
                underlined
                className="ms-Grid-col ms-sm12"
                name="justficationForTrip"
                value={reqData.justficationForTrip}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty(reqData.justficationForTrip),
                  "Justification for trip Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              />
            </Stack>
          </div>

          {/* Section B/C */}
          <br></br>
          <div className="ms-Grid">
            <div className="ms-Grid-Row">
              {/* Section B */}
              <div className="ms-Grid-col ms-sm6">
                <div className="ms-Grid-row">
                  <h2 className={styles.sectionHeader}>
                    Section B: Type of Travel (Select all that apply)
                  </h2>
                </div>
                {/*Conference Seminar Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        className={styles.generalLabel}
                        name="chbxConferenceSeminar"
                        id="chbxConferenceSeminar"
                        checked={reqData.chbxConferenceSeminar}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        Conference/Seminar**
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*Annual Auth. For Routine Travel Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxAnnualAuthForTravel"
                        id="chbxAnnualAuthForTravel"
                        checked={reqData.chbxAnnualAuthForTravel}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        Annual Auth. For Routine Travel
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*In-State Travel Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxInStateTravel"
                        id="chbxInStateTravel"
                        checked={reqData.chbxInStateTravel}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        In-State Travel
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*Out-Of-State Travel Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxOutOfStateTravel"
                        id="chbxOutOfStateTravel"
                        checked={reqData.chbxOutOfStateTravel}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        Out-Of-State Travel
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*Weekend Travel Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxWeekend"
                        id="chbxWeekend"
                        checked={reqData.chbxWeekend}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        Weekend Travel
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*Vehicle Rental Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxVehicleRental"
                        id="chbxVehicleRental"
                        checked={reqData.chbxVehicleRental}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        Vehicle Rental
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*Use Of Personal Vehicle Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxUserOfPersonalVehicle"
                        id="chbxUserOfPersonalVehicle"
                        checked={reqData.chbxUserOfPersonalVehicle}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        Use Of Personal Vehicle
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*Special Marketing Activity Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxSpecialMarketingActivities"
                        id="chbxSpecialMarketingActivities"
                        checked={reqData.chbxSpecialMarketingActivities}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        Special Marketing Activity
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*Prospect In The Same Hotel As Employee Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxProspectInSameHotelAsEmployee"
                        id="chbxProspectInSameHotelAsEmployee"
                        checked={reqData.chbxProspectInSameHotelAsEmployee}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        Prospect In The Same Hotel As Employee
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*50% Allowance above GSA Loding Rate Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbx50pctLodgingException"
                        id="chbx50pctLodgingException"
                        checked={reqData.chbx50pctLodgingException}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        50% Allowance above GSA Loding Rate
                      </label>
                    </Stack>
                  </div>
                </div>
                {/*Other (Please Attach Explanation) Checkbox*/}
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxOther"
                        id="chbxOther"
                        checked={reqData.chbxOther}
                        disabled={disableControls}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <label className={styles.generalLabel}>
                        Other (Please Attach Explanation)
                      </label>
                      <TextField
                        //className={styles.otherTextField}
                        underlined
                        name="chbxOtherSig"
                        value={reqData.chbxOtherSig}
                        disabled={disableControls}
                        onChange={this.handlereqDataTextChange.bind(this)}
                      />
                    </Stack>
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <p>
                    **REQUIRED DOCUMENTATION: If reason for trip is a Conference
                    or Seminar, a brochure or agenda is required to be attached
                    to this form.
                  </p>
                </div>
              </div>
              {/* Section C*/}
              <div className="ms-Grid-col ms-sm6">
                <div className="ms-Grid">
                  <div className="ms-Grid-row">
                    <h2 className={styles.sectionHeader}>
                      Section C: Estimated Expenses Per Traveler
                    </h2>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <label className={styles.sectionCLabel}>
                        Registration Fees:
                      </label>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          className={styles.SectionPriceColumn}
                          underlined
                          name="registrationFees"
                          value={reqData.registrationFees.toString()}
                          validateOnLoad={false}
                          //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                          disabled={disableControls}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "registrationFees"
                          )}
                          //onValueChange={this.handlereqDataNumberChange.bind(this, 'tipsAmount')} /> */}
                        />
                      </Stack>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <label className={styles.sectionCLabel}>
                        Airfare Costs:
                      </label>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          className={styles.SectionPriceColumn}
                          underlined
                          name="airFareCost"
                          value={reqData.airFareCost.toString()}
                          validateOnLoad={false}
                          //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                          disabled={disableControls}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "airFareCost"
                          )}
                        />
                      </Stack>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>
                          Personal Car Mileage:
                        </label>
                        <TextField
                          className={styles.SectionCTextbox}
                          underlined
                          name="mileageEstimation"
                          value={reqData.mileageEstimation}
                          disabled={disableControls}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "mileageEstimation"
                          )}
                        />
                        &nbsp;
                        <TextField
                          className={styles.SectionCTextboxHalfWidth}
                          underlined
                          name="mileageRate"
                          value={reqData.mileageRate.toString()}
                          validateOnLoad={false}
                          disabled={true}
                          onBlur={this.handlereqDataWholeNumberChange.bind(
                            this,
                            "mileageRate"
                          )}
                        />
                        <label className={styles.sectionCLabel}>??/mile:</label>
                      </Stack>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          className={styles.SectionPriceColumn}
                          underlined
                          name="mileageAmount"
                          value={reqData.mileageAmount.toString()}
                          validateOnLoad={false}
                          disabled={true}
                          onChange={this.handlereqDataNumberChange.bind(
                            this,
                            "mileageAmount"
                          )}
                        />
                      </Stack>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>
                          Lodging $
                        </label>
                        <TextField
                          //styles={{ root: { height: "20px" } }}
                          className={styles.SectionCTextbox}
                          underlined
                          name="lodgingCostPerNight"
                          value={reqData.lodgingCostPerNight}
                          disabled={disableControls}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "lodgingCostPerNight"
                          )}
                        />
                        <label className={styles.sectionCLabel}>x</label>
                        <TextField
                          //styles={{ root: { height: "20px" } }}
                          className={styles.SectionCTextbox}
                          underlined
                          name="lodgingNights"
                          value={reqData.lodgingNights}
                          disabled={disableControls}
                          onBlur={this.handlereqDataWholeNumberChange.bind(
                            this,
                            "lodgingNights"
                          )}
                        />
                        <label className={styles.sectionCLabel}>Nights:</label>
                      </Stack>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          className={styles.SectionPriceColumn}
                          underlined
                          name="totalLodgingAmount"
                          value={reqData.totalLodgingAmount.toString()}
                          validateOnLoad={false}
                          //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                          disabled={true}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "totalLodgingAmount"
                          )}
                        />
                      </Stack>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>Meals $</label>
                        <TextField
                          //styles={{ root: { height: "20px" } }}
                          className={styles.SectionCTextbox}
                          underlined
                          name="mealCostPerNight"
                          value={reqData.mealCostPerNight}
                          disabled={disableControls}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "mealCostPerNight"
                          )}
                        />
                        <label className={styles.sectionCLabel}>x</label>
                        <TextField
                          //styles={{ root: { height: "20px" } }}
                          className={styles.SectionCTextbox}
                          underlined
                          name="mealPerNights"
                          value={reqData.mealPerNights}
                          disabled={disableControls}
                          onBlur={this.handlereqDataWholeNumberChange.bind(
                            this,
                            "mealPerNights"
                          )}
                        />
                        <label className={styles.sectionCLabel}>Days:</label>
                      </Stack>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          className={styles.SectionPriceColumn}
                          underlined
                          name="totalMealAmount"
                          value={reqData.totalMealAmount.toString()}
                          validateOnLoad={false}
                          //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                          disabled={true}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "totalMealAmount"
                          )}
                        />
                      </Stack>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>
                          Car Rental:
                        </label>
                        &nbsp;&nbsp;&nbsp;
                        <Checkbox
                          name="carRentalUsed"
                          label="Yes"
                          id="chbxCarRentalYes"
                          checked={reqData.carRentalUsed == "true"}
                          onChange={this._onUniqueCheckboxChange.bind(
                            this,
                            "true"
                          )}
                          disabled={disableControls}
                          styles={checkboxStyles}
                        />
                        &nbsp;
                        <Checkbox
                          name="carRentalUsed"
                          label="No"
                          id="chbxCarRentalNo"
                          checked={reqData.carRentalUsed == "false"}
                          disabled={disableControls}
                          onChange={this._onUniqueCheckboxChange.bind(
                            this,
                            "false"
                          )}
                          styles={checkboxStyles}
                        />
                      </Stack>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          required={
                            reqData.carRentalUsed == "true" ? true : false
                          }
                          className={styles.SectionPriceColumn}
                          underlined
                          name="vehicleRentalCost"
                          value={reqData.vehicleRentalCost.toString()}
                          validateOnLoad={false}
                          //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                          disabled={disableControls}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "vehicleRentalCost"
                          )}
                        />
                      </Stack>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <label className={styles.sectionCLabel}>
                        Other Transport Costs (Taxi/Shuttle):
                      </label>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          className={styles.SectionPriceColumn}
                          underlined
                          name="otherTransportCosts"
                          value={reqData.otherTransportCosts.toString()}
                          validateOnLoad={false}
                          //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                          disabled={disableControls}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "otherTransportCosts"
                          )}
                        />
                      </Stack>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <label className={styles.sectionCLabel}>
                        Cost Per Traveler:
                      </label>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          className={styles.SectionPriceColumn}
                          underlined
                          name="costPerTraveler"
                          value={reqData.costPerTraveler.toString()}
                          validateOnLoad={false}
                          //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                          disabled={true}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "costPerTraveler"
                          )}
                        />
                      </Stack>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <label className={styles.sectionCLabel}>
                        Special Marketing Activity:
                      </label>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          className={styles.SectionPriceColumn}
                          underlined
                          name="specialMarketingActivitiesAmount"
                          value={reqData.specialMarketingActivitiesAmount.toString()}
                          validateOnLoad={false}
                          //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                          disabled={disableControls}
                          onBlur={this.handlereqDataNumberChange.bind(
                            this,
                            "specialMarketingActivitiesAmount"
                          )}
                        />
                      </Stack>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm9">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>
                          Number of Travelers:
                        </label>
                        &nbsp;
                        <TextField
                          //styles={{ root: { height: "20px" } }}
                          className={styles.SectionCTextboxHalfWidth}
                          underlined
                          name="numberOfTravelers"
                          value={reqData.numberOfTravelers}
                          disabled={disableControls}
                          onBlur={this.handlereqDataWholeNumberChange.bind(
                            this,
                            "numberOfTravelers"
                          )}
                        />
                        &nbsp;
                        <label className={styles.sectionCLabel}>Total:</label>
                      </Stack>
                    </div>
                    <div className="ms-Grid-col ms-sm3">
                      <Stack horizontal>
                        <label className={styles.sectionCLabel}>$</label>
                        <TextField
                          className={styles.SectionPriceColumn}
                          underlined
                          name="totalEstimatedCostOfTrip"
                          value={reqData.totalEstimatedCostOfTrip.toString()}
                          validateOnLoad={false}
                          //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                          disabled={true}
                          onChange={this.handlereqDataNumberChange.bind(
                            this,
                            "totalEstimatedCostOfTrip"
                          )}
                        />
                      </Stack>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Section D */}
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <h2 className={styles.sectionHeader}>
                Section D: Additional Travelers
              </h2>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6">
                <label className={styles.generalLabel}>Traveler Name</label>
                <TextField
                  underlined
                  name="TravelerName1"
                  value={reqData.TravelerName1}
                  required={false}
                  validateOnLoad={false}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="TravelerName2"
                  value={reqData.TravelerName2}
                  required={false}
                  validateOnLoad={false}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="TravelerName3"
                  value={reqData.TravelerName3}
                  required={false}
                  validateOnLoad={false}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
              <div className="ms-Grid-col ms-sm6">
                <label className={styles.generalLabel}>
                  Traveler Job Title
                </label>
                <TextField
                  underlined
                  name="TravelerjobTitle1"
                  value={reqData.TravelerjobTitle1}
                  required={false}
                  validateOnLoad={false}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="TravelerjobTitle2"
                  value={reqData.TravelerjobTitle2}
                  required={false}
                  validateOnLoad={false}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="TravelerjobTitle3"
                  value={reqData.TravelerjobTitle3}
                  required={false}
                  validateOnLoad={false}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
            </div>
          </div>

          {/* Section E */}
          <br></br>
          <br></br>
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <Stack horizontal>
                <h2 className={styles.sectionHeader}>
                  Section E: Agency Accounting
                </h2>
                <br></br>
                <TextField
                  className={styles.sectionETextbox}
                  underlined
                  name="agencyAccounting"
                  value={reqData.agencyAccounting}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                &nbsp;
                {/* <TextField
                  className={styles.sectionETextbox}
                  underlined
                  name="deputySecretary"
                  value={reqData.deputySecretary}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <h2 className={styles.sectionHeader}>Deputy Undersecretary</h2> */}
                {budget.approvalStatus !==
                  "Approved" /*&& budget.userLogin == currentUser.loginName*/ && (
                  <div>
                    <PrimaryButton
                      className="budgetButton"
                      data-automation-id="BudgetApprove"
                      text="Budget Approval"
                      title="budget"
                      disabled={!isBudgetApprover}
                      onClick={this.approvalButton.bind(this, "budget")}
                    />
                  </div>
                )}
              </Stack>
            </div>

            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm2">
                <label className={styles.generalLabel}>Agency</label>
                <TextField
                  underlined
                  name="Agency1"
                  value={reqData.Agency1}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="Agency2"
                  value={reqData.Agency2}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="Agency3"
                  value={reqData.Agency3}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
              <div className="ms-Grid-col ms-sm2">
                <label className={styles.generalLabel}>Cost Center</label>
                <TextField
                  underlined
                  name="CostCenter1"
                  value={reqData.CostCenter1}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="CostCenter2"
                  value={reqData.CostCenter2}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="CostCenter3"
                  value={reqData.CostCenter3}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
              <div className="ms-Grid-col ms-sm2">
                <label className={styles.generalLabel}>Fund</label>
                <TextField
                  underlined
                  name="Fund1"
                  value={reqData.Fund1}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="Fund2"
                  value={reqData.Fund2}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="Fund3"
                  value={reqData.Fund3}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
              <div className="ms-Grid-col ms-sm2">
                <label className={styles.generalLabel}>General Ledger</label>
                <TextField
                  underlined
                  name="GeneralLedger1"
                  value={reqData.GeneralLedger1}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="GeneralLedger2"
                  value={reqData.GeneralLedger2}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="GeneralLedger3"
                  value={reqData.GeneralLedger3}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
              <div className="ms-Grid-col ms-sm2">
                <label className={styles.generalLabel}>Grant #</label>
                <TextField
                  underlined
                  name="Grant1"
                  value={reqData.Grant1}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="Grant2"
                  value={reqData.Grant2}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="Grant3"
                  value={reqData.Grant3}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
              <div className="ms-Grid-col ms-sm2">
                <label className={styles.generalLabel}>WBS Element</label>
                <TextField
                  underlined
                  name="WBSElemenet1"
                  value={reqData.WBSElemenet1}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="WBSElemenet2"
                  value={reqData.WBSElemenet2}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <TextField
                  underlined
                  name="WBSElemenet3"
                  value={reqData.WBSElemenet3}
                  required={false}
                  validateOnLoad={false}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
            </div>
          </div>

          {/* Section F */}
          <br></br>
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <h2 className={styles.sectionHeader}>
                Section F: Approval Signature
              </h2>
            </div>
            {/* Employee Approval 
            <div className="ms-Grid-row">
              <div className={styles.approvalBox}>
                {" "}
                {reqData.employeeApproval.approvalStatus !== "Approved" &&
                  reqData.employeeApproval.userLogin !==
                    currentUser.loginName && (
                    <div className={styles.smallWhenPrinting}>
                      {" "}
                      Pending Approval from{" "}
                      {reqData.employeeApproval.displayName}
                    </div>
                  )}
                {reqData.employeeApproval.approvalStatus !== "Approved" &&
                  reqData.employeeApproval.userLogin ==
                    currentUser.loginName && (
                    <div>
                      <PrimaryButton
                        className={styles.buttonSpacing}
                        data-automation-id="employeeApproval"
                        disabled={disableSubmit}
                        text="Approve"
                        title="employeeApproval"
                        onClick={this.approvalButton.bind(
                          this,
                          "employeeApproval"
                        )}
                      />
                    </div>
                  )}
                {reqData.employeeApproval.approvalStatus == "Approved" && (
                  <div>
                    <div className={styles.smallWhenPrinting}>
                      {reqData.employeeApproval.approvalString}
                    </div>
                    {undersecretary.comment && (
                      <div>Comment: {undersecretary.comment}</div>
                    )}
                  </div>
                )}
              </div>
              <span className={styles.approvalTitle}>Employee</span>
            </div>
            <br></br>
            */}

            {/* Section Head Approval */}
            <div className="ms-Grid-row">
              <div className={styles.approvalBox}>
                {sectionHead.approvalStatus !== "Approved" &&
                  sectionHead.userLogin !== currentUser.loginName && (
                    <div className={styles.smallWhenPrinting}>
                      {sectionHead.approvalString}
                    </div>
                  )}
                {sectionHead.approvalStatus !== "Approved" &&
                  sectionHead.userLogin == currentUser.loginName && (
                    <div>
                      <PrimaryButton
                        className={styles.buttonSpacing}
                        data-automation-id="SectionHeadApprove"
                        text="Approve"
                        title="sectionHead"
                        onClick={this.approvalButton.bind(this, "sectionHead")}
                      />
                    </div>
                  )}
                {sectionHead.approvalStatus == "Approved" && (
                  <div>
                    <div className={styles.smallWhenPrinting}>
                      {sectionHead.approvalString}
                    </div>
                    {sectionHead.comment && (
                      <div>Comment: {sectionHead.comment}</div>
                    )}
                  </div>
                )}
              </div>
              <span className={styles.approvalTitle}>Section Head</span>
            </div>
            <br></br>

            {/* Department Head Approval */}
            <div className="ms-Grid-row">
              <div className={styles.approvalBox}>
                {secretary.approvalStatus !== "Approved" &&
                  secretary.userLogin !== currentUser.loginName && (
                    <div className={styles.smallWhenPrinting}>
                      {" "}
                      {secretary.approvalString}
                    </div>
                  )}
                {secretary.approvalStatus !== "Approved" &&
                  secretary.userLogin == currentUser.loginName && (
                    <div>
                      <PrimaryButton
                        className={styles.buttonSpacing}
                        data-automation-id="secretaryApprove"
                        text="Approve"
                        title="secretary"
                        disabled={disableSubmitForSpecialSigs}
                        onClick={this.approvalButton.bind(this, "secretary")}
                      />
                      {disableSubmitForSpecialSigs && (
                        <div>
                          Please ensure all Special Approvals are signed.
                        </div>
                      )}
                    </div>
                  )}
                {secretary.approvalStatus == "Approved" && (
                  <div>
                    <div className={styles.smallWhenPrinting}>
                      {secretary.approvalString}
                    </div>
                    {secretary.comment && (
                      <div>Comment: {secretary.comment}</div>
                    )}
                  </div>
                )}
              </div>
              <span className={styles.approvalTitle}>
                Department Head/Designee Signature
              </span>
            </div>
            <br></br>

            {/* Deputy Undersecretary Approval */}
            <div className="ms-Grid-row">
              <div className={styles.approvalBox}>
                {deputyUndersecretary.approvalStatus !== "Approved" &&
                  deputyUndersecretary.userLogin !== currentUser.loginName && (
                    <div className={styles.smallWhenPrinting}>
                      {deputyUndersecretary.approvalString}
                    </div>
                  )}
                {deputyUndersecretary.approvalStatus !== "Approved" &&
                  deputyUndersecretary.userLogin == currentUser.loginName && (
                    <div>
                      <PrimaryButton
                        className={styles.buttonSpacing}
                        data-automation-id="deputyUndersecretaryApprove"
                        text="Approve"
                        title="deputyUndersecretary"
                        disabled={disableSubmitForSpecialSigs}
                        onClick={this.approvalButton.bind(
                          this,
                          "deputyUndersecretary"
                        )}
                      />
                    </div>
                  )}
                {disableSubmitForSpecialSigs && (
                  <div>Please ensure all Special Approvals are signed.</div>
                )}
                {deputyUndersecretary.approvalStatus == "Approved" && (
                  <div>
                    <div className={styles.smallWhenPrinting}>
                      {deputyUndersecretary.approvalString}
                    </div>
                    {deputyUndersecretary.comment && (
                      <div>Comment: {deputyUndersecretary.comment}</div>
                    )}
                  </div>
                )}
              </div>
              <span className={styles.approvalTitle}>
                Deputy Undersecretary
              </span>
            </div>
          </div>
        </div>
        {/* Section G */}
        <br></br>
        <div>
          <h2 className={styles.sectionHeader}>Section G: Notes</h2>
          <TextField
            multiline
            underlined
            name="extraNotes"
            value={reqData.extraNotes}
            required={false}
            validateOnLoad={false}
            disabled={disableControls}
            onChange={this.handlereqDataTextChange.bind(this)}
          />
        </div>
        <br></br>

        {/*Section Test */}
        {/* <TextField
          underlined
          className="ms-Grid-col ms-sm6"
          name="employeeName"
          value={reqData.employeeApproval.displayName}
          validateOnLoad={false}
          onChange={this.handlereqDataTextChange.bind(this)}
        />
        <TextField
          underlined
          className="ms-Grid-col ms-sm6"
          name="employeeTitle"
          value={reqData.employeeApproval.jobTitle}
          validateOnLoad={false}
          onChange={this.handlereqDataTextChange.bind(this)}
        />
        <TextField
          underlined
          className="ms-Grid-col ms-sm6"
          name="employeeLogin"
          value={reqData.employeeApproval.userLogin}
          validateOnLoad={false}
          onChange={this.handlereqDataTextChange.bind(this)}
        />{" "} */}

        {/* Tool Options */}
        <br></br>
        <div>
          {reqData.status == "Draft" && (
            <PrimaryButton
              data-automation-id="test"
              disabled={disableSubmit}
              text="Submit"
              className={`${styles.buttonSpacing} ${styles.printHide}`}
              onClick={this.Submit.bind(this)}
            />
          )}
          <PrimaryButton
            onClick={this.SaveButton.bind(this)}
            text="Save"
            className={`${styles.buttonSpacing} ${styles.printHide}`}
          />
          <DefaultButton
            onClick={this.CloseForm.bind(this)}
            text="Close"
            className={`${styles.buttonSpacing} ${styles.printHide}`}
          />
          {/* <DefaultButton
            onClick={this.printPage.bind(this)}
            text="Print"
            className={`${styles.buttonSpacing} ${styles.printHide}`}
          /> */}
          <DefaultButton
            onClick={this.emailPDF.bind(this)}
            disabled={disableSubmit}
            text="Email PDF"
            className={`${styles.buttonSpacing} ${styles.printHide}`}
          />
          {this.state.saving == true && (
            <Spinner
              label="Saving Request..."
              ariaLive="assertive"
              labelPosition="right"
            />
          )}
        </div>
      </div>
    );
  }
}
