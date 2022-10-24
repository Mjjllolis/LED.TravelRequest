import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITravelRequestProps {
  mileageRate: number;
  context: WebPartContext;
  startingFinancialYear: number;
}
  