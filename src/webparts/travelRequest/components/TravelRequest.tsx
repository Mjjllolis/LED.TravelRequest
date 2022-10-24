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
        employeeId: null,
        employeeName: "",
        employeeLogin: "",
        agency: "",
        personnelNo: "",
        costCenter: "",
        domicile: "",
        taNo: "",
        departureTime: "",
        departureDateStr: "",
        returnTime: "",
        returnDateStr: "",
        fund: "",
        dateOfRequest: new Date(),
        fYBudget: "0.00",
        amtRemainBudget: "0.00",
        amtRemainingAfterThis: "0.00",
        authBudget: "0.00",
        gL: "",
        sMAGL: "",
        fySpecialMarketing: "0.00",
        fySpecialMarketingamtRemaining: "0.00",
        fySpecialMarketingamtRemainingAfterThis: "0.00",
        fYBudgetFY2: "0.00",
        amtRemainBudgetFY2: "0.00",
        amtRemainingAfterThisFY2: "0.00",
        authBudgetFY2: "0.00",
        fySpecialMarketingFY2: "0.00",
        fySpecialMarketingamtRemainingFY2: "0.00",
        fySpecialMarketingamtRemainingAfterThisFY2: "0.00",
        destination: "",
        status: "Draft",
        stage: "",
        nextApprover: null,
        requestLog: "",
        purposeOfTrip: "",
        benefitToState: "",
        airTravelAgencyUsed: null,
        airTravelAgencyUsedJustification: "",
        airFare: "",
        airFareCost: "0.00",
        vehicleType: "",
        mileageEstimation: 0.0,
        mileageRate: 0.0,
        mileageAmount: "0.00",
        vehiclePassengers: "",
        vehicleRentalTypeIsCompact: "",
        vehicleRentalJustificationChoice: "",
        vehicleRentalJustificationText: "",
        vehicleRentalCost: "0.00",
        limoTaxi: "",
        limoTaxiFareAmount: "0.00",
        tollsAndParking: "",
        tollsAndParkingAmount: "0.00",
        totalTransportationExpense: "0.00",
        lodging: [
          {
            total: 0.0,
            days: 0.0,
            cost: 0.0,
          },
          {
            total: 0.0,
            days: 0.0,
            cost: 0.0,
          },
          {
            total: 0.0,
            days: 0.0,
            cost: 0.0,
          },
        ],
        totalLodgingAmount: "0.00",
        meals: [
          {
            total: 0.0,
            days: 0.0,
            cost: 0.0,
          },
        ],
        totalMealAmount: "0.00",
        tips: "",
        tipsAmount: "0.00",
        otherExpensePayableTo: "",
        otherExpensePaymentMethod: "",
        otherExpenseDueDate: "",
        otherExpenseAmount: "0.00",
        totalEstimatedTravelAmount: "0.00",
        specialMarketingActivitiesAmountNotes: "",
        specialMarketingActivitiesAmount: "0.00",
        totalEstimatedCostOfTrip: "0.00",
        travelAdvanceDate: "",
        travelAdvanceAmount: "0.00",
        chbxVehicleRental: false,
        chbxGPSRentalVehicle: false,
        chbxProspectInSameHotelAsEmployee: false,
        chbxSpecialMarketingActivities: false,
        chbx50pctLodgingException: false,
        chbxOther: false,
        OtherExplanation: "",
        chbxVehicleRentalSig: "",
        chbxGPSRentalVehicleSig: "",
        chbxProspectInSameHotelAsEmployeeSig: "",
        chbxSpecialMarketingActivitiesSig: "",
        chbx50pctLodgingExceptionSig: "",
        chbxOtherSig: "",

        EstimatedCompensatoryTime: "",
        budgetYear1: 0,
        budgetYear2: 0,

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
  private handleMaskedreqDataDateChange(event) {
    const { name, value } = event.target;
    let reqData = { ...this.state.reqData };
    //let val = value.replace('_','');
    reqData[name] = value;
    reqData[name] = reqData[name].replace("_", "");
    this.setState({ reqData });
  }

  private async handleMaskedDateWithValidation(event) {
    const { name, value } = event.target;
    let ctrlName = name;
    let reqData = { ...this.state.reqData };
    reqData[ctrlName] = value.replace(/_/g, "");
    let combinedDateTime = new Date();
    let valiMessage = "Required";
    let needToValidate = false;
    let testDate = new Date();
    switch (ctrlName) {
      case "departureDateStr":
        combinedDateTime = new Date(
          reqData[ctrlName] + " " + reqData.departureTime
        );
        if (combinedDateTime.getTime() && reqData.departureTime) {
          reqData.departureDate = combinedDateTime;
          await this.setState({ DepartureDateError: "" });
        } else {
          valiMessage = "Departure Date and Time Required";
          await this.setState({
            DepartureDateError: "Valid Departure Date and Time Required",
          });
        }
        testDate = new Date(reqData[ctrlName]);
        if (!testDate.getTime() || !reqData[ctrlName]) {
          needToValidate = true;
        }
        break;

      case "departureTime":
        combinedDateTime = new Date(
          reqData.departureDateStr + " " + reqData[ctrlName]
        );
        if (combinedDateTime.getTime() && reqData.departureTime) {
          reqData.departureDate = combinedDateTime;
          await this.setState({ DepartureDateError: "" });
        } else {
          valiMessage = "Valid Departure Date and Time Required";
          await this.setState({
            DepartureDateError: "Valid Departure Date and Time Required",
          });
        }
        testDate = new Date("9/9/2009 " + reqData[ctrlName]);
        if (!testDate.getTime() || !reqData[ctrlName]) {
          needToValidate = true;
        }
        break;

      case "returnDateStr":
        combinedDateTime = new Date(
          reqData[ctrlName] + " " + reqData.returnTime
        );
        if (combinedDateTime.getTime() && reqData.returnTime) {
          reqData.returnDate = combinedDateTime;
          await this.setState({ ReturnDateError: "" });
        } else {
          valiMessage = "Valid Return Date and Time Required";
          await this.setState({
            ReturnDateError: "Valid Return Date and Time Required",
          });
        }
        testDate = new Date(reqData[ctrlName]);
        if (!testDate.getTime() || !reqData[ctrlName]) {
          needToValidate = true;
        }
        break;

      case "returnTime":
        combinedDateTime = new Date(
          reqData.returnDateStr + " " + reqData[ctrlName]
        );
        if (combinedDateTime.getTime() && reqData.returnTime) {
          reqData.returnDate = combinedDateTime;
          await this.setState({ ReturnDateError: "" });
        } else {
          valiMessage = "Valid Return Date and Time Required";
          await this.setState({
            ReturnDateError: "Valid Return Date and Time Required",
          });
        }
        testDate = new Date("9/9/2009 " + reqData[ctrlName]);
        if (!testDate.getTime() || !reqData[ctrlName]) {
          needToValidate = true;
        }
        break;
    }
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

  private _onControlledCheckboxChange(event) {
    const { name, checked } = event.target;
    let reqData = { ...this.state.reqData };
    reqData[name] = checked;
    this.setState({ reqData });
  }
  private async _onUniqueCheckboxChange(checkboxVal, event) {
    const { name } = event.target;
    let reqData = { ...this.state.reqData };
    checkboxVal = reqData[name] == checkboxVal ? "" : checkboxVal; //to allow for unchecking
    reqData[name] = checkboxVal;
    await this.setState({ reqData });
    this.updateCurrencyCalculations(this);
  }

  private handlereqDataRadioChange(event, option: any) {
    const { name } = event.target;
    const val = option.key;
    let reqData = { ...this.state.reqData };
    reqData[name] = val;
    this.setState({ reqData });
  }

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

  private async handlereqDataNumberChangeOLD(fieldName, value) {
    let reqData = { ...this.state.reqData };
    let val = !isNaN(value.floatValue) ? value.floatValue : "";
    reqData[fieldName] = val;
    await this.setState({ reqData });
    this.updateCurrencyCalculations(this);
  }

  private async handleMultiDayNumberChange(arrayName, index, prop, value) {
    let reqData = { ...this.state.reqData };
    let val = !isNaN(value.floatValue) ? value.floatValue : null;
    reqData[arrayName][index][prop] = val;
    await this.setState({ reqData });
    this.updateCurrencyCalculations(this);
  }
  private async _addMultiDay(arrayName, event) {
    let reqData = { ...this.state.reqData };
    let newMDay = new MultidayCost();
    reqData[arrayName].push({ total: 0, days: 0, cost: 0 });
    await this.setState({ reqData });
    this.updateCurrencyCalculations(this);
  }
  private async _removeMultiDay(arrayName, index, event) {
    let reqData = { ...this.state.reqData };
    reqData[arrayName].splice(index, 1);
    await this.setState({ reqData });
    this.updateCurrencyCalculations(this);
  }

  private printPage() {
    window.print();
  }

  private updateCurrencyCalculations(ctx) {
    let reqData = { ...this.state.reqData };

    //demo calculation, not sure if it's needed, probably not correct
    if (
      Number(reqData.amtRemainBudget.replace(/,/g, "")) > 0 &&
      Number(reqData.authBudget.replace(/,/g, "")) > 0
    ) {
      reqData.amtRemainingAfterThis = (
        Number(reqData.amtRemainBudget.replace(/,/g, "")) -
        Number(reqData.authBudget.replace(/,/g, ""))
      )
        .toFixed(2)
        .replace(/\d(?=(\d{3})+\.)/g, "$&,");
    }
    //trim trailing decimals
    let miles = reqData.mileageEstimation ? reqData.mileageEstimation : 0.0;
    let airFare = reqData.airFareCost
      ? Number(reqData.airFareCost.replace(/,/g, ""))
      : 0.0;
    let vehicleRentalCost = reqData.vehicleRentalCost
      ? Number(reqData.vehicleRentalCost.replace(/,/g, ""))
      : 0.0;
    let limoTaxiFareAmount = reqData.limoTaxiFareAmount
      ? Number(reqData.limoTaxiFareAmount.replace(/,/g, ""))
      : 0.0;
    let mileageRate = reqData.mileageRate ? Number(reqData.mileageRate) : 0.0;
    reqData.mileageAmount =
      reqData.vehicleType == "Personal"
        ? (miles * mileageRate).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, "$&,")
        : "0.00";
    reqData.totalTransportationExpense = (
      airFare +
      Number(reqData.mileageAmount.replace(/,/g, "")) +
      vehicleRentalCost +
      limoTaxiFareAmount
    )
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");

    //add lodging costs
    let tempCost = 0.0;
    for (const lodge of reqData.lodging) {
      lodge.total = lodge.days && lodge.cost ? lodge.days * lodge.cost : 0.0;
      tempCost = tempCost + lodge.total;
    }
    reqData.totalLodgingAmount = tempCost
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");

    //add meals costs
    tempCost = 0.0;
    for (const meal of reqData.meals) {
      meal.total = meal.days && meal.cost ? meal.days * meal.cost : 0.0;
      tempCost = tempCost + meal.total;
    }
    reqData.totalMealAmount = tempCost
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");

    //TOTALS
    let tollsAndParkingAmount = reqData.tollsAndParkingAmount
      ? Number(reqData.tollsAndParkingAmount.replace(/,/g, ""))
      : 0.0;
    let tipsAmount = reqData.tipsAmount
      ? Number(reqData.tipsAmount.replace(/,/g, ""))
      : 0.0;
    let otherExpenseAmount = reqData.otherExpenseAmount
      ? Number(reqData.otherExpenseAmount.replace(/,/g, ""))
      : 0.0;
    reqData.totalEstimatedTravelAmount = (
      Number(reqData.totalTransportationExpense.replace(/,/g, "")) +
      Number(reqData.totalLodgingAmount.replace(/,/g, "")) +
      Number(reqData.totalMealAmount.replace(/,/g, "")) +
      tollsAndParkingAmount +
      tipsAmount +
      otherExpenseAmount
    )
      .toFixed(2)
      .replace(/\d(?=(\d{3})+\.)/g, "$&,");

    //total Estimated Cost of trip
    let specialMarketingActivitiesAmount =
      reqData.specialMarketingActivitiesAmount
        ? Number(reqData.specialMarketingActivitiesAmount.replace(/,/g, ""))
        : 0.0;
    reqData.totalEstimatedCostOfTrip = (
      Number(reqData.totalEstimatedTravelAmount.replace(/,/g, "")) +
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

  private requiredNumberValidation(value) {
    return isNaN(value.floatValue) ? false : true;
  }

  //when people picker changes, update state
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

  private _onSelectDate(id, date: Date | null | undefined) {
    let reqData = { ...this.state.reqData };
    reqData[id] = date;
    this.setState({ reqData });
  }

  private _onSelectDD(id, event) {
    if (event.type == "click") {
      const { innerText } = event.target;
      let reqData = { ...this.state.reqData };
      reqData[id] = innerText;
      this.setState({ reqData });
    }
  }

  private _onFormatDate = (date: Date): string => {
    return date.toLocaleDateString();
  };
  private handleCommentChange(event) {
    const { name, value } = event.target;
    var st = { ...this.state };
    st.reqData[name].comment = value;
    this.setState(st);
  }
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  };

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
        //Employee
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
          st.reqData.stage = "Secretary";
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
        //Section Head
        if (
          st.reqData.secretary.approvalStatus == "" &&
          st.reqData.secretary.userId
        ) {
          st.reqData.stage = "Secretary";
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

    let specApprArray = [
      "chbxVehicleRental",
      "chbxGPSRentalVehicle",
      "chbxProspectInSameHotelAsEmployee",
      "chbxSpecialMarketingActivities",
      "chbx50pctLodgingException",
      "chbxOther",
    ];
    let unapprovedSpecs = false;
    specApprArray.forEach((e) => {
      if (
        st.reqData[e] == true &&
        st.reqData[e + "Sig"] == "" &&
        (approvalName == "deputyUndersecretary" ||
          approvalName == "undersecretary")
      ) {
        unapprovedSpecs = true;
      }
    });
    if (unapprovedSpecs) {
      //toast.success("One or more special approvals still need approval! Please Approve and Save Form");
      toast.warn(
        "One or items in the Special Approvals Section still need approval! Please Approve and Save Form"
      );
    } else {
      toast.success("Form approved!");
    }
  }
  private rejectButton(approval: Approver, event) {
    //check and only continue if comment is added

    //prompt to ensure that user wants to cancel existing approvals and restart process

    //remove approval status on all other approvals and set stage

    //set rejection info with user name, date and comments to be used in emails and logs

    //append rejection info to request log

    //save form
    this.SaveButton();
  }

  private async SaveButton() {
    this.setState({ saving: true });
    let itemId = await this.service.SaveRequest(this.state);
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
    if (this.state.reqData.status == "Draft") {
      let reqData = { ...this.state.reqData };
      reqData.status = "In Progress";
      let kickoffFlowValue = "Yes";
      this.setState({ kickoffFLOW: "Yes" });
      if (
        reqData.employeeApproval.userLogin ==
          this.props.context.pageContext.user.loginName.toLowerCase() &&
        reqData.employeeApproval.approvalStatus == ""
      ) {
        this.approvalButton("employeeApproval");
      } else if (reqData.employeeApproval.approvalStatus == "") {
        reqData.stage = "Employee Approval";
        reqData.nextApprover = reqData.employeeApproval.userId;
        await this.setState({ reqData });
        await this.service.SaveRequest(this.state, kickoffFlowValue);
      }
    } else {
      await this.service.SaveRequest(this.state);
    }
    this.CloseForm();
  }

  private CloseForm() {
    let queryParameters = new UrlQueryParameterCollection(window.location.href);
    const srcUrl = queryParameters.getValue("Source");
    let webUrl = this.props.context.pageContext.web.absoluteUrl;
    window.location.href = srcUrl ? srcUrl : webUrl;
  }

  //handle adding attachments
  private showForm() {
    //let reqData = { ...this.state.reqData };
    //reqData.Adding= true;
    //this.setState({ reqData });
    this.setState({ AddingAttachment: true });
  }

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

  private async RemoveAttachment(attachment) {
    await this.service.RemoveAttachment(attachment.Id);
    setTimeout(
      function () {
        this.FetchAttachments();
      }.bind(this),
      1000
    );
  }

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
      data.budgetYear1 = data.budgetYear1
        ? data.budgetYear1
        : Number(this.props.startingFinancialYear.toString().slice(-2));
      data.budgetYear2 = data.budgetYear2
        ? data.budgetYear2
        : Number(this.props.startingFinancialYear.toString().slice(-2)) + 1;
      this.setState({ reqData: data });
      //this._addMultiDay('meals', null);
      //this._addMultiDay('lodging', null);
    }

    this._getAndSetApprovers();

    let reqData = { ...this.state.reqData };
    //sectionHead.userLogin = "admin@laecondev.onmicrosoft.com";
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
    let disableSubmitForSpecialSigs = false;
    if (
      !reqData.agency ||
      !reqData.personnelNo ||
      !reqData.departureDateStr ||
      !reqData.returnDateStr ||
      !reqData.destination ||
      !reqData.purposeOfTrip ||
      !reqData.benefitToState ||
      !reqData.domicile
    ) {
      disableSubmit = true;
    }
    if (
      (reqData.stage == "Secretary" ||
        reqData.stage == "Undersecretary" ||
        reqData.stage == "Deputy Undersecretary") &&
      ((reqData.chbxVehicleRental && !reqData.chbxVehicleRentalSig) ||
        (reqData.chbxGPSRentalVehicle && !reqData.chbxGPSRentalVehicleSig) ||
        (reqData.chbx50pctLodgingException &&
          !reqData.chbx50pctLodgingExceptionSig) ||
        (reqData.chbxOther && !reqData.chbxOtherSig) ||
        (reqData.chbxProspectInSameHotelAsEmployee &&
          !reqData.chbxProspectInSameHotelAsEmployeeSig) ||
        (reqData.chbxSpecialMarketingActivities &&
          !reqData.chbxSpecialMarketingActivitiesSig))
    ) {
      disableSubmitForSpecialSigs = true;
    }

    const isBudgetApprover =
      budget.userLogin == currentUser.loginName.toLowerCase() ? true : false;
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
          <div className="container">
            <div className="ms-Grid-row">
              <h2 className="col align-self-center title">
                Travel Request Form
              </h2>
            </div>
            <div className="ms-Grid-row">
              <p>
                NO REGISTRATIONS OR RESERVATIONS SHOULD BE MADE UNTIL ALL
                APPROVALS ARE OBTAINED Instructions: Complete all sections
                pertaining to your request. Print or Type all entries. Submit
                completed form with all necessary approvals to your Agency’s
                Travel Administrator. Retain a copy for your records.
              </p>
            </div>
          </div>

          {/* Section A Row 1*/}
          <div className="container">
            <div className="ms-Grid-row">
              <h2 className="ms-Grid-col">
                Section A: General Information- Complete All Info
              </h2>
            </div>
            <div className="ms-Grid-row">
              <TextField
                className="ms-Grid-col"
                underlined
                label="Name:"
                name="name"
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
              />
              <TextField
                className="ms-Grid-col"
                underlined
                label="Destination:"
                name="Destination"
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
              <Stack horizontal>
                <Label required>Departure:</Label>
                <MaskedInput
                  mask="11/11/1111"
                  name="departureDateStr"
                  onChange={this.handleMaskedDateWithValidation.bind(this)}
                  value={reqData.departureDateStr}
                  className={styles.inputMaskControl}
                  disabled={disableControls}
                  required={true}
                />
                <Label required>Return:</Label>
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
            </div>
          </div>

          {/* Section A Row 2*/}
          <div className="container">
            <div className="ms-Grid-row">
              <TextField
                className="ms-Grid-col"
                label="Agency:"
                name="Agency"
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
              <TextField
                className="ms-Grid-col"
                label="Section:"
                name="Section"
                value={"Section"}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty("Section"),
                  "Section Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              />
              <TextField
                className="ms-Grid-col"
                label="Mode of Transportation:"
                name="Mode of Transportation"
                value={"Mode Of Transportation"}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty("Test"),
                  "Mode Of Transportation Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              />
            </div>
          </div>

          {/* Section A Row 3*/}
          <div className="container">
            <div className="ms-Grid-row">
              <TextField
                label="Purpose/Justification For Travel:"
                name="Purpose/Justification For Travel"
                value={"Purpose Of Trip"}
                required={true}
                validateOnLoad={false}
                onGetErrorMessage={this.genericValidation.bind(
                  this,
                  name,
                  stringIsNullOrEmpty("Purpose Of Trip"),
                  "Purpose For Trip Required"
                )}
                disabled={disableControls}
                onChange={this.handlereqDataTextChange.bind(this)}
              />
            </div>
          </div>
        </div>

        {/* Old */}
        <br></br>
        <br></br>
        <br></br>
        {/* Old */}

        <div className="form-group">
          <div className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md7 ms-lg7">
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                    {/* <Label className={styles.label+' '+styles.printHide}>Employee:</Label> */}
                    <Label className={styles.label}>Employee:</Label>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                    <div className={styles.printHide}>
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
                    <div className={styles.printShow}>
                      {reqData.employeeName}
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg4">
                    <TextField
                      label="Agency:"
                      underlined
                      name="agency"
                      id=""
                      value={reqData.agency}
                      required={true}
                      disabled={disableControls}
                      validateOnLoad={false}
                      onGetErrorMessage={this.genericValidation.bind(
                        this,
                        name,
                        stringIsNullOrEmpty(reqData.agency),
                        "Agency Required"
                      )}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                    <TextField
                      label="Personnel #:"
                      underlined
                      name="personnelNo"
                      value={reqData.personnelNo}
                      validateOnLoad={false}
                      onGetErrorMessage={this.genericValidation.bind(
                        this,
                        name,
                        stringIsNullOrEmpty(reqData.personnelNo),
                        "Personnel Number Required"
                      )}
                      onChange={this.handlereqDataTextChange.bind(this)}
                      disabled={disableControls}
                      required
                    />
                    <TextField
                      label="TA #:"
                      underlined
                      readOnly
                      name="taNo"
                      value={reqData.taNo}
                      //required={true}
                      validateOnLoad={false}
                      //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.taNo), 'TA Number Required')}
                      disabled={true}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />

                    <div className={styles.underlineText}>
                      <Stack horizontal>
                        <Label required>Departure:</Label>
                        <MaskedInput
                          mask="11/11/1111"
                          name="departureDateStr"
                          onChange={this.handleMaskedDateWithValidation.bind(
                            this
                          )}
                          value={reqData.departureDateStr}
                          className={styles.inputMaskControl}
                          disabled={disableControls}
                          required={true}
                        />
                        <Label required>Time:</Label>
                        <MaskedInput
                          mask="11:11 aa"
                          name="departureTime"
                          onChange={this.handleMaskedDateWithValidation.bind(
                            this
                          )}
                          value={reqData.departureTime}
                          className={styles.inputMaskControl}
                          disabled={disableControls}
                          required
                        />
                      </Stack>
                    </div>
                    {this.state.DepartureDateError && (
                      <div className={styles.validationMessage}>
                        {this.state.DepartureDateError}
                      </div>
                    )}
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
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
                      onSelectDate={this._onSelectDate.bind(
                        this,
                        "dateOfRequest"
                      )}
                    />
                    <TextField
                      label="Official Domicile:"
                      underlined
                      name="domicile"
                      value={reqData.domicile}
                      required={true}
                      validateOnLoad={false}
                      onGetErrorMessage={this.genericValidation.bind(
                        this,
                        name,
                        stringIsNullOrEmpty(reqData.domicile),
                        "Official domicile Required"
                      )}
                      disabled={disableControls}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />

                    <div className={styles.underlineText}>
                      <Stack horizontal>
                        <Label required>Return:</Label>
                        <MaskedInput
                          mask="11/11/1111"
                          name="returnDateStr"
                          onChange={this.handleMaskedDateWithValidation.bind(
                            this
                          )}
                          value={reqData.returnDateStr}
                          className={styles.inputMaskControl}
                          disabled={disableControls}
                          required={true}
                        />
                        <Label required>Time:</Label>
                        <MaskedInput
                          mask="11:11 aa"
                          name="returnTime"
                          onChange={this.handleMaskedDateWithValidation.bind(
                            this
                          )}
                          value={reqData.returnTime}
                          className={styles.inputMaskControl}
                          disabled={disableControls}
                          required
                        />
                      </Stack>
                    </div>
                    {this.state.ReturnDateError && (
                      <div className={styles.validationMessage}>
                        {this.state.ReturnDateError}
                      </div>
                    )}
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <TextField
                    label="Destination:"
                    underlined
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
                </div>
                <div className="ms-Grid-row">
                  <TextField
                    label="Purpose of Trip:"
                    underlined
                    //multiline
                    autoAdjustHeight
                    name="purposeOfTrip"
                    value={reqData.purposeOfTrip}
                    required={true}
                    validateOnLoad={false}
                    onGetErrorMessage={this.genericValidation.bind(
                      this,
                      name,
                      stringIsNullOrEmpty(reqData.purposeOfTrip),
                      "Answer Required"
                    )}
                    disabled={disableControls}
                    onChange={this.handlereqDataTextChange.bind(this)}
                  />
                </div>
                <div className="ms-Grid-row">
                  <TextField
                    label="Benefit to State:"
                    description="(Explanation required for all out-of-state travel, except for prospecting/missions)"
                    underlined
                    name="benefitToState"
                    value={reqData.benefitToState}
                    required={true}
                    validateOnLoad={false}
                    onGetErrorMessage={this.genericValidation.bind(
                      this,
                      name,
                      stringIsNullOrEmpty(reqData.benefitToState),
                      "Answer Required"
                    )}
                    disabled={disableControls}
                    onChange={this.handlereqDataTextChange.bind(this)}
                  />
                </div>
                <div className="ms-Grid-row">
                  <Label className={"airTravelQuestion"}>
                    {" "}
                    Air Travel Arranged Through Contracted Travel Agency:
                  </Label>
                  <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                    <Checkbox
                      name="airTravelAgencyUsed"
                      label="Yes"
                      id="airTravelAgencyYes"
                      checked={reqData.airTravelAgencyUsed == "true"}
                      onChange={this._onUniqueCheckboxChange.bind(this, "true")}
                      disabled={disableControls}
                      styles={checkboxStyles}
                    />
                    <Checkbox
                      name="airTravelAgencyUsed"
                      label="No"
                      id="airTravelAgencyNo"
                      checked={reqData.airTravelAgencyUsed == "false"}
                      disabled={disableControls}
                      onChange={this._onUniqueCheckboxChange.bind(
                        this,
                        "false"
                      )}
                      styles={checkboxStyles}
                    />
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                    <TextField
                      label="Explain:"
                      name="airTravelAgencyUsedJustification"
                      value={reqData.airTravelAgencyUsedJustification}
                      //multiline
                      autoAdjustHeight
                      required={
                        reqData.airTravelAgencyUsed == "false" ? true : false
                      }
                      validateOnLoad={false}
                      onGetErrorMessage={this.genericValidation.bind(
                        this,
                        name,
                        stringIsNullOrEmpty(
                          reqData.airTravelAgencyUsedJustification
                        ),
                        "Answer Required"
                      )}
                      disabled={disableControls}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />
                  </div>
                </div>
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md5 ms-lg5">
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    {/* <h2>
                      <label >Budget Section </label>
                    </h2> */}
                    <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                      <TextField
                        label="CC:"
                        underlined
                        name="costCenter"
                        value={reqData.costCenter}
                        validateOnLoad={false}
                        onGetErrorMessage={this.genericValidation.bind(
                          this,
                          name,
                          stringIsNullOrEmpty(reqData.costCenter),
                          "Cost Center Required"
                        )}
                        disabled={!isBudgetApprover}
                        onChange={this.handlereqDataTextChange.bind(this)}
                      />
                    </div>
                    <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                      <TextField
                        label="Fund:"
                        underlined
                        name="fund"
                        value={reqData.fund}
                        validateOnLoad={false}
                        onGetErrorMessage={this.genericValidation.bind(
                          this,
                          name,
                          stringIsNullOrEmpty(reqData.fund),
                          "Fund Required"
                        )}
                        disabled={!isBudgetApprover}
                        onChange={this.handlereqDataTextChange.bind(this)}
                      />
                    </div>
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                      <TextField
                        label="GL:"
                        underlined
                        name="gL"
                        value={reqData.gL}
                        validateOnLoad={false}
                        onGetErrorMessage={this.genericValidation.bind(
                          this,
                          name,
                          stringIsNullOrEmpty(reqData.fund),
                          "GL Required"
                        )}
                        disabled={!isBudgetApprover}
                        onChange={this.handlereqDataTextChange.bind(this)}
                      />
                    </div>
                    <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                      <TextField
                        label="SMAL:" //label="Special Marketing Activities Ledger:"
                        underlined
                        name="sMAGL"
                        value={reqData.sMAGL}
                        validateOnLoad={false}
                        onGetErrorMessage={this.genericValidation.bind(
                          this,
                          name,
                          stringIsNullOrEmpty(reqData.fund),
                          "Special Marketing Activities Ledger Required"
                        )}
                        disabled={!isBudgetApprover}
                        onChange={this.handlereqDataTextChange.bind(this)}
                      />
                    </div>
                  </div>
                </div>

                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                    <h3>
                      <label>Budget Year </label>
                    </h3>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <h3>
                      <label>{`FY${reqData.budgetYear1}`}</label>
                    </h3>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <h3>
                      <label>{`FY${reqData.budgetYear2}`}</label>
                    </h3>
                  </div>
                </div>

                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                    <div className={styles.underlineText}>
                      FY Travel Budget:
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    {/* <div className={`${styles.currencyFldWrapper}  ${styles.addDollarSign}`}> */}
                    <div className={`${styles.currencyFldWrapper} `}>
                      {/* <CurrencyFormat
                        thousandSeparator={true}

                        value={reqData.fYBudget ? reqData.fYBudget : ""}
                        //isAllowed={this.requiredNumberValidation.bind(this)}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'fYBudget')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.fYBudget}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "fYBudget"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.fYBudgetFY2}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "fYBudgetFY2"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}

                        value={reqData.fYBudgetFY2 ? reqData.fYBudgetFY2 : ""}
                        //isAllowed={this.requiredNumberValidation.bind(this)}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'fYBudgetFY2')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                </div>

                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                    <div className={styles.underlineText}>
                      Amt. Remaining in Budget:
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.amtRemainBudget}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "amtRemainBudget"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}

                        value={reqData.amtRemainBudget ? reqData.amtRemainBudget : ""}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'amtRemainBudget')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.amtRemainBudgetFY2}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "amtRemainBudgetFY2"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}

                        value={reqData.amtRemainBudgetFY2 ? reqData.amtRemainBudgetFY2 : ""}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'amtRemainBudgetFY2')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                </div>

                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                    <div className={styles.underlineText}>
                      Amt. after Authorization:
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.amtRemainingAfterThis}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "amtRemainingAfterThis"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}

                        value={reqData.amtRemainingAfterThis ? reqData.amtRemainingAfterThis : ""}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'amtRemainingAfterThis')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.amtRemainingAfterThisFY2}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "amtRemainingAfterThisFY2"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}

                        value={reqData.amtRemainingAfterThisFY2 ? reqData.amtRemainingAfterThisFY2 : ""}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'amtRemainingAfterThisFY2')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                </div>

                <hr />

                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                    <div className={styles.underlineText}>
                      FY Special Marketing Activities:
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.fySpecialMarketing}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "fySpecialMarketing"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}
                        value={reqData.fySpecialMarketing ? reqData.fySpecialMarketing : ""}
                        //isAllowed={this.requiredNumberValidation.bind(this)}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'fySpecialMarketing')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.fySpecialMarketingFY2}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "fySpecialMarketingFY2"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}
                        value={reqData.fySpecialMarketingFY2 ? reqData.fySpecialMarketingFY2 : ""}
                        //isAllowed={this.requiredNumberValidation.bind(this)}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'fySpecialMarketingFY2')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                </div>

                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                    <div className={styles.underlineText}>
                      Amt. Remaining in Budget:
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.fySpecialMarketingamtRemaining}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "fySpecialMarketingamtRemaining"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}
                        value={reqData.fySpecialMarketingamtRemaining ? reqData.fySpecialMarketingamtRemaining : ""}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'fySpecialMarketingamtRemaining')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.fySpecialMarketingamtRemainingFY2}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "fySpecialMarketingamtRemainingFY2"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}
                        value={reqData.fySpecialMarketingamtRemainingFY2 ? reqData.fySpecialMarketingamtRemainingFY2 : ""}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'fySpecialMarketingamtRemainingFY2')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                </div>

                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm4 ms-md6 ms-lg6">
                    <div className={styles.underlineText}>
                      Amt. after Authorization:
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={reqData.fySpecialMarketingamtRemainingAfterThis}
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "fySpecialMarketingamtRemainingAfterThis"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}
                        value={reqData.fySpecialMarketingamtRemainingAfterThis ? reqData.fySpecialMarketingamtRemainingAfterThis : ""}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'fySpecialMarketingamtRemainingAfterThis')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm4 ms-md3 ms-lg3">
                    <div className={`${styles.currencyFldWrapper} `}>
                      <CurrencyTextField
                        //label="Amount"
                        variant="standard"
                        value={
                          reqData.fySpecialMarketingamtRemainingAfterThisFY2
                        }
                        //currencySymbol="$"
                        outputFormat="number"
                        onBlur={this.handlereqDataNumberChange.bind(
                          this,
                          "fySpecialMarketingamtRemainingAfterThisFY2"
                        )}
                        disabled={!isBudgetApprover}
                        className={styles.currencyFormatting}
                      />
                      {/* <CurrencyFormat
                        thousandSeparator={true}
                        value={reqData.fySpecialMarketingamtRemainingAfterThisFY2 ? reqData.fySpecialMarketingamtRemainingAfterThisFY2 : ""}
                        disabled={!isBudgetApprover}
                        onValueChange={this.handlereqDataNumberChange.bind(this, 'fySpecialMarketingamtRemainingAfterThisFY2')}
                        decimalScale={2}
                        fixedDecimalScale={true}
                      /> */}
                    </div>
                  </div>
                </div>

                {budget.approvalStatus !==
                  "Approved" /*&& budget.userLogin == currentUser.loginName*/ && (
                  <div>
                    <PrimaryButton
                      className={styles.buttonSpacing}
                      data-automation-id="BudgetApprove"
                      text="Budget Amounts Added"
                      title="budget"
                      disabled={!isBudgetApprover}
                      onClick={this.approvalButton.bind(this, "budget")}
                    />
                  </div>
                )}
              </div>
            </div>

            <hr />

            <div className="ms-Grid-row">
              <div className="ms-Grid-col  ms-sm12 ms-md8 ms-lg8">
                <TextField
                  label="Air Fare (Coach Class):"
                  underlined
                  name="airFare"
                  value={reqData.airFare}
                  //required={true}
                  //validateOnLoad={false}
                  //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.airFare), 'Answer Required')}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.airFareCost}
                    //currencySymbol="$"
                    outputFormat="number"
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "airFareCost"
                    )}
                    disabled={disableControls}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat
                    thousandSeparator={true}
                    value={reqData.airFareCost ? reqData.airFareCost : ""}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    //placeholder='$0.00'
                    //isAllowed={this.requiredNumberValidation.bind(this)}
                    disabled={disableControls}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'airFareCost')}
                  /> */}
                </div>
              </div>
            </div>

            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <Label required={true}>Vehicle Used:</Label>
                <Checkbox
                  name="vehicleType"
                  label="State Car"
                  id="vehicleStateCar"
                  checked={reqData.vehicleType == "State"}
                  onChange={this._onUniqueCheckboxChange.bind(this, "State")}
                  disabled={disableControls}
                  styles={checkboxStyles}
                />
                <Checkbox
                  name="vehicleType"
                  label="Personal Car"
                  id="PersonalStateCar"
                  checked={reqData.vehicleType == "Personal"}
                  onChange={this._onUniqueCheckboxChange.bind(this, "Personal")}
                  disabled={disableControls}
                  styles={checkboxStyles}
                />
              </div>
              <Label
                required={reqData.vehicleType == "Personal" ? true : false}
              >
                If Personal Car, Indicate Estimated Mileage:
              </Label>
              <div className="ms-Grid-col ms-sm12 ms-md3 ms-lg3">
                <Stack horizontal>
                  <div className={styles.currencyFldWrapper}>
                    <CurrencyFormat
                      placeholder={"0"}
                      value={
                        reqData.mileageEstimation
                          ? reqData.mileageEstimation
                          : ""
                      }
                      suffix=""
                      required={
                        reqData.vehicleType == "Personal" ? true : false
                      }
                      //displayType={ reqData.vehicleRentalType == 'personal' ? 'input' : 'text'}
                      validateOnLoad={false}
                      //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                      disabled={disableControls}
                      onValueChange={this.handlereqDataNumberChangeOLD.bind(
                        this,
                        "mileageEstimation"
                      )}
                    />
                  </div>
                  <div className={` ${styles.padTopAndSides}`}>Miles at</div>
                </Stack>
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md3 ms-lg3">
                <Stack horizontal>
                  <div className={styles.currencyFldWrapper}>
                    <CurrencyFormat
                      name="mileageRate"
                      value={reqData.mileageRate ? reqData.mileageRate : ""}
                      suffix={""}
                      displayType={!isAcctMgr ? "text" : "input"}
                      validateOnLoad={false}
                      //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                      disabled={disableControls}
                      onValueChange={this.handlereqDataNumberChangeOLD.bind(
                        this,
                        "mileageRate"
                      )}
                    />
                  </div>
                  <div className={` ${styles.padTopAndSides}`}>¢ per Mile</div>
                </Stack>
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.mileageAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onChange={this.handlereqDataNumberChange.bind(
                      this,
                      "mileageAmount"
                    )}
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "mileageAmount"
                    )}
                    disabled={true}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat
                    thousandSeparator={true}
                    value={reqData.mileageAmount ? reqData.mileageAmount : ""}
                    placeholder='$0.00'
                    displayType='text'
                    //isAllowed={this.requiredNumberValidation.bind(this)}
                    disabled={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'mileageAmount')}
                    decimalScale={2}
                    fixedDecimalScale={true}
                  /> */}
                </div>
              </div>
            </div>

            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md10 ms-lg10">
                <TextField
                  label="List Passengers' Names, if not a State Employee:"
                  name="vehiclePassengers"
                  value={reqData.vehiclePassengers}
                  //multiline
                  autoAdjustHeight
                  underlined
                  validateOnLoad={false}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
            </div>

            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg4">
                <Label>Rental Car:</Label>
                <Checkbox
                  name="vehicleRentalJustificationChoice"
                  label="Compact/Subcompact"
                  id="Compact"
                  checked={
                    reqData.vehicleRentalJustificationChoice ==
                    "Compact/Subcompact"
                  }
                  onChange={this._onUniqueCheckboxChange.bind(
                    this,
                    "Compact/Subcompact"
                  )}
                  disabled={disableControls}
                  styles={checkboxStyles}
                />
                <Checkbox
                  name="vehicleRentalJustificationChoice"
                  label="No Compact/Subcompact Available"
                  id="NoCompactAvailable"
                  checked={
                    reqData.vehicleRentalJustificationChoice == "None Available"
                  }
                  onChange={this._onUniqueCheckboxChange.bind(
                    this,
                    "None Available"
                  )}
                  disabled={disableControls}
                  styles={checkboxStyles}
                />
                <Checkbox
                  name="vehicleRentalJustificationChoice"
                  label="Transporting more than 2 persons"
                  id="TransportingOver2"
                  checked={
                    reqData.vehicleRentalJustificationChoice ==
                    "Multiple Passengers"
                  }
                  onChange={this._onUniqueCheckboxChange.bind(
                    this,
                    "Multiple Passengers"
                  )}
                  disabled={disableControls}
                  styles={checkboxStyles}
                />
                <Checkbox
                  name="vehicleRentalJustificationChoice"
                  label="Same cost as Compact/Subcompact"
                  id="SameCost"
                  checked={
                    reqData.vehicleRentalJustificationChoice == "Equal Cost"
                  }
                  onChange={this._onUniqueCheckboxChange.bind(
                    this,
                    "Equal Cost"
                  )}
                  disabled={disableControls}
                  styles={checkboxStyles}
                />
                <Checkbox
                  name="vehicleRentalJustificationChoice"
                  label="Other: (Explain)"
                  id="RentalJustOther"
                  checked={reqData.vehicleRentalJustificationChoice == "Other"}
                  onChange={this._onUniqueCheckboxChange.bind(this, "Other")}
                  disabled={disableControls}
                  styles={checkboxStyles}
                />
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg4">
                <TextField
                  label="Explain why it is cost effective to use a rental:"
                  name="vehicleRentalJustificationText"
                  value={reqData.vehicleRentalJustificationText}
                  //multiline
                  autoAdjustHeight
                  validateOnLoad={false}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.vehicleRentalCost}
                    //currencySymbol="$"
                    outputFormat="number"
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "vehicleRentalCost"
                    )}
                    disabled={disableControls}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Cost:'
                    value={reqData.vehicleRentalCost ? reqData.vehicleRentalCost : ""}
                    thousandSeparator={true}
                    //required={ reqData.vehicleType=='Limo, Taxi, etc.' ? true : false}
                    //displayType={ reqData.vehicleRentalType == 'personal' ? 'input' : 'text'}
                    validateOnLoad={false}
                    //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                    disabled={disableControls}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'vehicleRentalCost')} /> */}
                </div>
              </div>
            </div>

            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8">
                <TextField
                  label="LIMOUSINE, TAXI, ETC."
                  underlined
                  name="limoTaxi"
                  value={reqData.limoTaxi}
                  //required={true}
                  //validateOnLoad={false}
                  //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.limoTaxi), 'Answer Required')}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.limoTaxiFareAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "limoTaxiFareAmount"
                    )}
                    disabled={disableControls}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Fare Amount:'
                    value={reqData.limoTaxiFareAmount ? reqData.limoTaxiFareAmount : ""}
                    thousandSeparator={true}
                    //required={ reqData.vehicleType=='Limo, Taxi, etc.' ? true : false}
                    //displayType={ reqData.vehicleRentalType == 'personal' ? 'input' : 'text'}
                    validateOnLoad={false}
                    //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                    disabled={disableControls}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'limoTaxiFareAmount')} /> */}
                </div>
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.totalTransportationExpense}
                    //currencySymbol="$"
                    outputFormat="number"
                    onChange={this.handlereqDataNumberChange.bind(
                      this,
                      "totalTransportationExpense"
                    )}
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "totalTransportationExpense"
                    )}
                    disabled={true}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Fare Amount:'
                    displayType='text'
                    thousandSeparator={true}
                    value={reqData.totalTransportationExpense ? reqData.totalTransportationExpense : ""}
                    //required={ reqData.vehicleType=='Limo, Taxi, etc.' ? true : false}
                    //displayType={ reqData.vehicleRentalType == 'personal' ? 'input' : 'text'}
                    disabled={disableControls}
                    validateOnLoad={false}
                    //onGetErrorMessage={this.genericValidation.bind(this, name, stringIsNullOrEmpty(reqData.mileageEstimation), 'Answer Required')}
                    //onChange={this.handlereqDataNumberChange.bind(this)} 
                    decimalScale={2}
                    fixedDecimalScale={true}
                  /> */}
                </div>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <Label>SUBSISTENCE</Label>
                {/* <ActionButton iconProps={addIcon} allowDisabledFocus disabled={disableControls} onClick={this._addMultiDay.bind(this, 'lodging')}>
                  Add Lodging
                </ActionButton> */}
              </div>
              <div className={`ms-Grid-col ms-sm10 ms-md8 ms-lg8`}>
                {reqData.lodging.map((lodge, i) => (
                  <div key={i}>
                    <Stack horizontal wrap className="smallCurrency">
                      <div className={`${styles.padTopAndSides}`}>Lodging:</div>
                      <div>
                        <div
                          className={`${styles.currencyFldWrapper} ${styles.subsistenceInputs}`}
                        >
                          <CurrencyFormat
                            label="Lodging:"
                            value={lodge.days ? lodge.days : ""}
                            validateOnLoad={false}
                            disabled={disableControls}
                            onValueChange={this.handleMultiDayNumberChange.bind(
                              this,
                              "lodging",
                              i,
                              "days"
                            )}
                          />
                        </div>
                      </div>
                      <div className={` ${styles.padTopAndSides}`}>
                        Nights @
                      </div>
                      <div>
                        <div
                          className={`${styles.currencyFldWrapper} ${styles.addDollarSign}`}
                        >
                          <CurrencyFormat
                            label="Lodging:"
                            thousandSeparator={true}
                            value={lodge.cost ? lodge.cost : ""}
                            prefix=""
                            suffix=""
                            validateOnLoad={false}
                            disabled={disableControls}
                            onValueChange={this.handleMultiDayNumberChange.bind(
                              this,
                              "lodging",
                              i,
                              "cost"
                            )}
                          />
                        </div>
                      </div>
                      <div className={` ${styles.padTopAndSides}`}>
                        /night Total:
                      </div>
                      <div>
                        <div
                          className={`${styles.currencyFldWrapper} ${styles.addDollarSign}`}
                        >
                          <CurrencyFormat
                            label="Lodging:"
                            displayType="text"
                            thousandSeparator={true}
                            placeholder="$0.00"
                            value={lodge.total ? lodge.total : ""}
                            disabled={disableControls}
                            validateOnLoad={false}
                            decimalScale={2}
                            fixedDecimalScale={true}
                          />
                        </div>
                      </div>
                      {/* {i !== 0 &&
                        <ActionButton iconProps={removeIcon} allowDisabledFocus onClick={this._removeMultiDay.bind(this, 'lodging', i)}>
                        </ActionButton>
                      } */}
                    </Stack>
                  </div>
                ))}
              </div>
              <Stack
                verticalAlign="end"
                className={`ms-Grid-col ms-sm2 ms-md2 ms-lg2 smallCurrency`}
              >
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.totalLodgingAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onChange={this.handlereqDataNumberChange.bind(
                      this,
                      "totalLodgingAmount"
                    )}
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "totalLodgingAmount"
                    )}
                    disabled={true}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Total Lodging:'
                    displayType='text'
                    thousandSeparator={true}
                    value={reqData.totalLodgingAmount ? reqData.totalLodgingAmount : ""}
                    disabled={disableControls}
                    validateOnLoad={false}
                    decimalScale={2}
                    fixedDecimalScale={true}
                  /> */}
                </div>
              </Stack>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <Label>MEALS</Label>
                {/* <ActionButton iconProps={addIcon} allowDisabledFocus disabled={disableControls} onClick={this._addMultiDay.bind(this, 'meals')}>
                  Add Meal
                </ActionButton> */}
              </div>
              <div className={`ms-Grid-col ms-sm10 ms-md8 ms-lg8`}>
                {reqData.meals.map((meal, i) => (
                  <div key={i}>
                    <Stack horizontal className="smallCurrency">
                      <div className={`${styles.padTopAndSides}`}>Meals:</div>
                      <div>
                        <div
                          className={`${styles.currencyFldWrapper} ${styles.subsistenceInputs}`}
                        >
                          <CurrencyFormat
                            label="Meal:"
                            value={meal.days ? meal.days : ""}
                            validateOnLoad={false}
                            disabled={disableControls}
                            onValueChange={this.handleMultiDayNumberChange.bind(
                              this,
                              "meals",
                              i,
                              "days"
                            )}
                          />
                        </div>
                      </div>
                      <div className={` ${styles.padTopAndSides}`}>Days @</div>
                      <div>
                        <div
                          className={`${styles.currencyFldWrapper} ${styles.addDollarSign}`}
                        >
                          <CurrencyFormat
                            label="Meal:"
                            value={meal.cost ? meal.cost : ""}
                            thousandSeparator={true}
                            prefix=""
                            suffix=""
                            validateOnLoad={false}
                            disabled={disableControls}
                            onValueChange={this.handleMultiDayNumberChange.bind(
                              this,
                              "meals",
                              i,
                              "cost"
                            )}
                          />
                        </div>
                      </div>
                      <div className={` ${styles.padTopAndSides}`}>
                        /day Total:
                      </div>
                      <div>
                        <div
                          className={`${styles.currencyFldWrapper} ${styles.addDollarSign}`}
                        >
                          <CurrencyFormat
                            label="Meal:"
                            displayType="text"
                            thousandSeparator={true}
                            placeholder="$0.00"
                            disabled={disableControls}
                            value={meal.total ? meal.total : ""}
                            validateOnLoad={false}
                            decimalScale={2}
                            fixedDecimalScale={true}
                          />
                        </div>
                      </div>
                      {/* {i !== 0 &&
                        <ActionButton iconProps={removeIcon} allowDisabledFocus onClick={this._removeMultiDay.bind(this, 'meals', i)}>
                        </ActionButton>
                      } */}
                    </Stack>
                  </div>
                ))}
              </div>
              <Stack
                verticalAlign="end"
                className={`ms-Grid-col ms-sm2 ms-md2 ms-lg2 smallCurrency`}
              >
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.totalMealAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onChange={this.handlereqDataNumberChange.bind(
                      this,
                      "totalMealAmount"
                    )}
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "totalMealAmount"
                    )}
                    disabled={true}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Total Meals:'
                    displayType='text'
                    thousandSeparator={true}
                    disabled={disableControls}
                    value={reqData.totalMealAmount ? reqData.totalMealAmount : ""}
                    validateOnLoad={false}
                    decimalScale={2}
                    fixedDecimalScale={true}
                  /> */}
                </div>
              </Stack>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md10 ms-lg10">
                <TextField
                  label="TOLLS AND PARKING"
                  underlined
                  name="tollsAndParking"
                  value={reqData.tollsAndParking}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.tollsAndParkingAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "tollsAndParkingAmount"
                    )}
                    disabled={disableControls}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Toll and Parking Amount:'
                    value={reqData.tollsAndParkingAmount ? reqData.tollsAndParkingAmount : ""}
                    thousandSeparator={true}
                    validateOnLoad={false}
                    disabled={disableControls}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'tollsAndParkingAmount')} /> */}
                </div>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md10 ms-lg10">
                <TextField
                  label="TIPS"
                  underlined
                  name="tips"
                  value={reqData.tips}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.tipsAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "tipsAmount"
                    )}
                    disabled={disableControls}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Tips Amount:'
                    value={reqData.tipsAmount ? reqData.tipsAmount : ""}
                    thousandSeparator={true}
                    validateOnLoad={false}
                    disabled={disableControls}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'tipsAmount')} /> */}
                </div>
              </div>
            </div>

            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md3 ms-lg3">
                <Label>REGISTRATION/OTHER</Label>
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md7 ms-lg7">
                <TextField
                  label="Payable To:"
                  underlined
                  name="otherExpensePayableTo"
                  value={reqData.otherExpensePayableTo}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md5 ms-lg5">
                    <div className={styles.underlineText}>
                      Due Date:
                      <MaskedInput
                        mask="11/11/1111"
                        name="otherExpenseDueDate"
                        onChange={this.handleMaskedreqDataDateChange.bind(this)}
                        value={reqData.otherExpenseDueDate}
                        className={`${styles.inputMaskControl} ${styles.inputMaskDateOnly}`}
                        disabled={disableControls}
                      />
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md7 ms-lg7">
                    <TextField
                      label="Payment Method:"
                      underlined
                      name="otherExpensePaymentMethod"
                      value={reqData.otherExpensePaymentMethod}
                      disabled={disableControls}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />
                  </div>
                </div>
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.otherExpenseAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "otherExpenseAmount"
                    )}
                    disabled={disableControls}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Registration Cost:'
                    value={reqData.otherExpenseAmount ? reqData.otherExpenseAmount : ""}
                    thousandSeparator={true}
                    validateOnLoad={false}
                    disabled={disableControls}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'otherExpenseAmount')} /> */}
                </div>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className={`ms-Grid-col ms-sm2 ms-md10 ms-lg10`}>
                <div className={styles.underlineText}>
                  TOTAL ESTIMATED TRAVEL COST:
                </div>
              </div>
              <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.totalEstimatedTravelAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onChange={this.handlereqDataNumberChange.bind(
                      this,
                      "totalEstimatedTravelAmount"
                    )}
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "totalEstimatedTravelAmount"
                    )}
                    disabled={true}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Total Estimated Travel Cost:'
                    value={reqData.totalEstimatedTravelAmount ? reqData.totalEstimatedTravelAmount : ""}
                    thousandSeparator={true}
                    displayType='text'
                    validateOnLoad={false}
                    disabled={disableControls}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'totalEstimatedTravelAmount')} /> */}
                </div>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm2 ms-md10 ms-lg10">
                <TextField
                  label="SPECIAL MARKETING ACTIVITIES:"
                  underlined
                  name="specialMarketingActivitiesAmountNotes"
                  value={reqData.specialMarketingActivitiesAmountNotes}
                  disabled={disableControls}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />
              </div>
              <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.specialMarketingActivitiesAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "specialMarketingActivitiesAmount"
                    )}
                    disabled={disableControls}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Special Marketing Activities:'
                    value={reqData.specialMarketingActivitiesAmount ? reqData.specialMarketingActivitiesAmount : ""}
                    thousandSeparator={true}
                    validateOnLoad={false}
                    disabled={disableControls}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'specialMarketingActivitiesAmount')} /> */}
                </div>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className={`ms-Grid-col ms-sm2 ms-md10 ms-lg10 `}>
                <div className={styles.underlineText}>
                  TOTAL ESTIMATED COST OF TRIP:
                </div>
              </div>
              <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.totalEstimatedCostOfTrip}
                    //currencySymbol="$"
                    outputFormat="number"
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "totalEstimatedCostOfTrip"
                    )}
                    disabled={true}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Total Estimated Travel Cost:'
                    value={reqData.totalEstimatedCostOfTrip ? reqData.totalEstimatedCostOfTrip : ""}
                    thousandSeparator={true}
                    validateOnLoad={false}
                    displayType='text'
                    disabled={disableControls}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'totalEstimatedCostOfTrip')} /> */}
                </div>
              </div>
            </div>

            <div className={`ms-Grid-row ${styles.GreyedOut}`}>
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={styles.underlineText}>TRAVEL ADVANCE</div>
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8">
                <div className={styles.underlineText}>
                  Date Needed:
                  <MaskedInput
                    mask="11/11/1111"
                    name="travelAdvanceDate"
                    onChange={this.handleMaskedreqDataDateChange.bind(this)}
                    value={reqData.travelAdvanceDate}
                    className={styles.inputMaskControl}
                    disabled={disableControls}
                  />
                </div>
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                <div className={`${styles.currencyFldWrapper} `}>
                  <CurrencyTextField
                    //label="Amount"
                    variant="standard"
                    value={reqData.travelAdvanceAmount}
                    //currencySymbol="$"
                    outputFormat="number"
                    onBlur={this.handlereqDataNumberChange.bind(
                      this,
                      "travelAdvanceAmount"
                    )}
                    disabled={disableControls}
                    className={styles.currencyFormatting}
                  />
                  {/* <CurrencyFormat label='Travel Advance Amount:'
                    value={reqData.travelAdvanceAmount ? reqData.travelAdvanceAmount : ""}
                    thousandSeparator={true}
                    validateOnLoad={false}
                    disabled={disableControls}
                    decimalScale={2}
                    fixedDecimalScale={true}
                    onValueChange={this.handlereqDataNumberChange.bind(this, 'travelAdvanceAmount')} /> */}
                </div>
              </div>
            </div>
          </div>

          <hr />
          <div className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row">
              <label className={styles.smallWhenPrinting}>
                I Hereby Certify That I Have A Valid LA Drivers License and When
                Applicable, I Futher Certify That I Have Vehicular Liability
                Insurance On My Personal Auto
              </label>
              {/* <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                <h2>Approval Section</h2>
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
              </div> */}
            </div>
          </div>

          <div className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                <h3>Special Approvals Required</h3>

                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxVehicleRental"
                        label=""
                        id="specApprVehicleRental"
                        checked={reqData.chbxVehicleRental}
                        //disabled={ !isApprover }
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        styles={checkboxStyles}
                      />
                      <div className={styles.specialApprText}>
                        Vehicle Rental
                      </div>
                    </Stack>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                    <TextField
                      name="chbxVehicleRentalSig"
                      //styles={specialApproverSigStyles}
                      underlined
                      value={reqData.chbxVehicleRentalSig}
                      disabled={!isApprover}
                      required={
                        reqData.chbxVehicleRental &&
                        reqData.stage == "Secretary"
                      }
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxGPSRentalVehicle"
                        label=""
                        id="specApprGps"
                        checked={reqData.chbxGPSRentalVehicle}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        //disabled={ !isApprover }
                        styles={checkboxStyles}
                      />
                      <div className={styles.specialApprText}>
                        GPS/Rental Vehicle
                      </div>
                    </Stack>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                    <TextField
                      name="chbxGPSRentalVehicleSig"
                      underlined
                      value={reqData.chbxGPSRentalVehicleSig}
                      disabled={!isApprover}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxProspectInSameHotelAsEmployee"
                        label=""
                        id="specApprSameHotel"
                        checked={reqData.chbxProspectInSameHotelAsEmployee}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        //disabled={ !isApprover }
                        styles={checkboxStyles}
                      />
                      <div className={styles.specialApprText}>
                        Prospect in Same Hotel as Employee
                      </div>
                    </Stack>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                    <TextField
                      name="chbxProspectInSameHotelAsEmployeeSig"
                      underlined
                      value={reqData.chbxProspectInSameHotelAsEmployeeSig}
                      disabled={!isApprover}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxSpecialMarketingActivities"
                        label=""
                        id="specApprSpecMarketing"
                        checked={reqData.chbxSpecialMarketingActivities}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        //disabled={ !isApprover }
                        styles={checkboxStyles}
                      />
                      <div className={styles.specialApprText}>
                        Special Marketing Activities
                      </div>
                    </Stack>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                    <TextField
                      name="chbxSpecialMarketingActivitiesSig"
                      underlined
                      value={reqData.chbxSpecialMarketingActivitiesSig}
                      disabled={!isApprover}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8">
                    <Stack horizontal>
                      <Checkbox
                        name="chbx50pctLodgingException"
                        label=""
                        id="specAppr50LodgingExc"
                        checked={reqData.chbx50pctLodgingException}
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        //disabled={ !isApprover }
                        styles={checkboxStyles}
                      />
                      <div className={styles.specialApprText}>
                        50% Lodging Exception
                      </div>
                    </Stack>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                    <TextField
                      name="chbx50pctLodgingExceptionSig"
                      underlined
                      value={reqData.chbx50pctLodgingExceptionSig}
                      disabled={!isApprover}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md8 ms-lg8">
                    <Stack horizontal>
                      <Checkbox
                        name="chbxOther"
                        checked={reqData.chbxOther}
                        id="specApprOther"
                        onChange={this._onControlledCheckboxChange.bind(this)}
                        //disabled={ !isApprover }
                        styles={checkboxStyles}
                      />
                      <div className={styles.specialApprText}>Other</div>
                      <TextField
                        //label="Other"
                        className={styles.widthInHorizontalStackFix}
                        //description="(Please Explain)"
                        underlined
                        name="OtherExplanation"
                        value={reqData.OtherExplanation}
                        //disabled={ !isApprover }
                        onChange={this.handlereqDataTextChange.bind(this)}
                      />
                    </Stack>
                  </div>
                  <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                    <TextField
                      name="chbxOtherSig"
                      underlined
                      value={reqData.chbxOtherSig}
                      disabled={!isApprover}
                      onChange={this.handlereqDataTextChange.bind(this)}
                    />
                  </div>
                </div>
                <br />

                <TextField
                  label="Notes"
                  underlined
                  autoAdjustHeight
                  //multiline
                  //multiline
                  //multiline
                />
                <TextField
                  label="Estimated Compensatory Time (See Attached)"
                  name="EstimatedCompensatoryTime"
                  underlined
                  value={reqData.EstimatedCompensatoryTime}
                  onChange={this.handlereqDataTextChange.bind(this)}
                />

                <h5 className={styles.smallWhenPrinting}>
                  When Prospecting, those requirements involving auto rental and
                  50% lodging exception will be considered to have met the
                  documentation requirements of PPM 49 when approved by the
                  agency head.
                </h5>

                <div className={styles.printHide}>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                      <h2>
                        <label>Attachments </label>
                      </h2>
                      <div>
                        <div className="ms-Grid">
                          {this.state.Attachments.map((att, i) => (
                            <div className="ms-Grid-row">
                              <Stack horizontal>
                                <a
                                  href={`${this.props.context.pageContext.web.absoluteUrl}/FormAttachments/Forms/AllItems.aspx?id=${this.props.context.pageContext.web.serverRelativeUrl}/FormAttachments/${att["FileLeafRef"]}&parent=${this.props.context.pageContext.web.serverRelativeUrl}/FormAttachments`}
                                  target="_blank"
                                >
                                  <div className={styles.attachText}>
                                    {att["FileLeafRef"]}
                                  </div>
                                </a>
                                <div>
                                  <ActionButton
                                    iconProps={{ iconName: "RemoveFilter" }}
                                    onClick={this.RemoveAttachment.bind(
                                      this,
                                      att
                                    )}
                                  ></ActionButton>
                                </div>
                              </Stack>
                            </div>
                          ))}
                        </div>
                      </div>

                      <ActionButton
                        className={styles.actionBtn}
                        iconProps={{ iconName: "PageAdd" }}
                        onClick={this.showForm.bind(this)}
                      >
                        New Attachment
                      </ActionButton>
                    </div>
                  </div>
                </div>
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                <h3>Approvals</h3>
                <div className="ms-Grid" dir="ltr">
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2">
                      <div className={styles.smallWhenPrinting}>
                        Employee Signature:
                      </div>
                    </div>
                    <div className={`ms-Grid-col ms-sm10 ms-md10 ms-lg10 `}>
                      <div className={styles.approvalBox}>
                        {" "}
                        {reqData.employeeApproval.approvalStatus !==
                          "Approved" &&
                          reqData.employeeApproval.userLogin !==
                            currentUser.loginName && (
                            <div className={styles.smallWhenPrinting}>
                              {" "}
                              Pending Approval from{" "}
                              {reqData.employeeApproval.displayName}
                            </div>
                          )}
                        {reqData.employeeApproval.approvalStatus !==
                          "Approved" &&
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
                        {reqData.employeeApproval.approvalStatus ==
                          "Approved" && (
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
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2">
                      <div className={styles.smallWhenPrinting}>Status:</div>
                    </div>
                    <div className={`ms-Grid-col ms-sm10 ms-md10 ms-lg10 `}>
                      <div className={styles.approvalBox}>
                        {" "}
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
                                onClick={this.approvalButton.bind(
                                  this,
                                  "sectionHead"
                                )}
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
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2">
                      <div className={styles.smallWhenPrinting}>Status:</div>
                    </div>
                    <div className={`ms-Grid-col ms-sm10 ms-md10 ms-lg10 `}>
                      <div className={styles.approvalBox}>
                        {" "}
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
                                onClick={this.approvalButton.bind(
                                  this,
                                  "secretary"
                                )}
                              />
                              {disableSubmitForSpecialSigs && (
                                <div>
                                  Please ensure all Special Approvals are
                                  signed.
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
                        Department/Deputy/Assistant Secretary
                      </span>
                    </div>
                  </div>
                  <h5 className={styles.smallWhenPrinting}>
                    I Certify That This Voucher Has Been Examined, That The
                    Proposed Expenditure Is Authorized By Appropriation and
                    Allotment And Does Not Exceed The Estimated Balance Of the
                    Allotment To Which It Is Properly Chargeable.
                  </h5>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2">
                      <div className={styles.smallWhenPrinting}>Status:</div>
                    </div>
                    <div className={`ms-Grid-col ms-sm10 ms-md10 ms-lg10 `}>
                      <div className={styles.approvalBox}>
                        {" "}
                        {undersecretary.approvalStatus !== "Approved" &&
                          undersecretary.userLogin !==
                            currentUser.loginName && (
                            <div className={styles.smallWhenPrinting}>
                              {undersecretary.approvalString}
                            </div>
                          )}
                        {undersecretary.approvalStatus !== "Approved" &&
                          undersecretary.userLogin == currentUser.loginName && (
                            <div>
                              <PrimaryButton
                                className={styles.buttonSpacing}
                                data-automation-id="undersecretaryApprove"
                                text="Approve"
                                title="undersecretary"
                                disabled={disableSubmitForSpecialSigs}
                                onClick={this.approvalButton.bind(
                                  this,
                                  "undersecretary"
                                )}
                              />
                            </div>
                          )}
                        {disableSubmitForSpecialSigs && (
                          <div>
                            Please ensure all Special Approvals are signed.
                          </div>
                        )}
                        {undersecretary.approvalStatus == "Approved" && (
                          <div>
                            <div className={styles.smallWhenPrinting}>
                              {undersecretary.approvalString}
                            </div>
                            {undersecretary.comment && (
                              <div>Comment: {undersecretary.comment}</div>
                            )}
                          </div>
                        )}
                      </div>
                      <span className={styles.approvalTitle}>
                        Undersecretary
                      </span>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2">
                      <div className={styles.smallWhenPrinting}>Status:</div>
                    </div>
                    <div className={`ms-Grid-col ms-sm10 ms-md10 ms-lg10 `}>
                      <div className={styles.approvalBox}>
                        {deputyUndersecretary.approvalStatus !== "Approved" &&
                          deputyUndersecretary.userLogin !==
                            currentUser.loginName && (
                            <div className={styles.smallWhenPrinting}>
                              {deputyUndersecretary.approvalString}
                            </div>
                          )}
                        {deputyUndersecretary.approvalStatus !== "Approved" &&
                          deputyUndersecretary.userLogin ==
                            currentUser.loginName && (
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
                          <div>
                            Please ensure all Special Approvals are signed.
                          </div>
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
              </div>

              <Dialog
                hidden={this.state.hideDialog}
                onDismiss={this._closeDialog}
                dialogContentProps={{
                  type: DialogType.normal,
                  title: this.state.dialogTitle,
                  subText: this.state.dialogText,
                }}
                modalProps={{
                  isBlocking: false,
                  styles: { main: { maxWidth: 450 } },
                }}
              >
                <DialogFooter>
                  {this.state.dialogTitle == "Approval" && (
                    <div>
                      <PrimaryButton
                        onClick={this.SaveAndCloseButton.bind(this)}
                        text="Save & Close"
                        className={styles.buttonSpacing}
                      />
                      <DefaultButton
                        onClick={this.SaveButton.bind(this)}
                        text="Save & Continue"
                      />
                    </div>
                  )}
                </DialogFooter>
              </Dialog>
            </div>
          </div>
        </div>
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
          <DefaultButton
            onClick={this.printPage.bind(this)}
            text="Print"
            className={`${styles.buttonSpacing} ${styles.printHide}`}
          />
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
        <AddAttachment
          isOpen={AddingAttachment}
          context={this.props.context}
          onClose={this._onClose.bind(this)}
          formKey={reqData.formKey}
        />
      </div>
    );
  }
}
