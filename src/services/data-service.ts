import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PageContext } from "@microsoft/sp-page-context";
import { sp, Web, ItemAddResult } from "@pnp/sp";
import { IAddAttachmentState } from "../webparts/travelRequest/components/AddAttachment";
import { stringIsNullOrEmpty, getRandomString } from "@pnp/common";
import "@pnp/polyfill-ie11";

export class DataService {
  private pc: PageContext;
  constructor(context: PageContext) {
    this.pc = context;
  }

  public async getRequestData(requestId: string) {
    return new Promise<any[]>(async (resolve, reject) => {
      sp.web.lists
        .getByTitle("Travel Requests")
        .items.select(
          "Title",
          "Status",
          "Stage",
          "NextApprover",
          "RequestLog",
          "RequestData"
        )
        .orderBy("Title", true)
        .getById(Number(requestId))
        .get()
        .then((res) => {
          resolve(res);
        })
        .catch((e) => {
          reject(e);
        });
    });
  }

  public SaveRequest(st, kickoffValue?) {
    return new Promise<string>(async (resolve, reject) => {
      let req = st.reqData;
      let data = {
        EmployeeId: req.employeeId, //have to add "Id to the end of the internal name"
        RequestDate: JSON.stringify(req.dateOfRequest)
          .replace('"', "")
          .replace('"', ""),
        Status: req.status,
        Stage: req.stage,
        NextApproverId: req.nextApprover,
        RequestLog: req.requestLog,
        RequestData: JSON.stringify(req),
        AuthorizationBudget: req.authBudget,
        kickoffFLOW: kickoffValue,
      };
      //if (st.formMode !== 'Edit') {
      let itemId = null;
      let dateObj = new Date(req.departureDate);
      let dateval =
        req.departureDate && dateObj.getTime()
          ? dateObj.toLocaleDateString("en-US")
          : "";
      let title = `${req.employeeName} - ${dateval} [${req.destination}]`;
      if (!st.requestID) {
        try {
          // add an item to the list
          data["Title"] = title;
          let item = await sp.web.lists
            .getByTitle("Travel Requests")
            .items.add(data);
          console.log(item);
          itemId = item.data.Id;
        } catch (e) {
          reject(e);
          console.log("Error adding form item");
        }
      } else {
        itemId = st.requestID;
      }
      data["Request"] = {
        __metadata: { type: "SP.FieldUrlValue" },
        Description: title,
        Url: `${
          this.pc.web.serverRelativeUrl
        }/SitePages/Request.aspx?RequestID=${itemId.toString()}`,
      };
      sp.web.lists
        .getByTitle("Travel Requests")
        .items.getById(itemId)
        .update(data)
        .then(() => {
          resolve(itemId);
        })
        .catch((e) => {
          reject(e);
        });
    });
  }

  public SaveEmailSubmission(id) {
    return new Promise<string>(async (resolve, reject) => {
      // let req = st.reqData;
      let data = {
        Title: "Sample Title",
        TravelReqID: id,
      };
      //if (st.formMode !== 'Edit') {
      let itemId = null;
      // let dateObj = new Date(req.departureDate)
      // let dateval = req.departureDate && dateObj.getTime() ? dateObj.toLocaleDateString('en-US') : "";
      // let title = `${req.employeeName} - ${dateval} [${req.destination}]`;
      // if(!st.requestID){
      try {
        // add an item to the list
        // data['Title'] = title;
        let item = await sp.web.lists
          .getByTitle("PrintRequests")
          .items.add(data);
        console.log(item);
        itemId = item.data.Id;
      } catch (e) {
        reject(e);
        console.log("Error adding form item");
      }
      // }
      // else{
      //     itemId = st.requestID;
      // }
      // data['Request'] = {
      //     "__metadata": { type: "SP.FieldUrlValue" },
      //     Description: title,
      //     Url: `${this.pc.web.serverRelativeUrl}/SitePages/Request.aspx?RequestID=${itemId.toString()}`
      // };
      // sp.web.lists.getByTitle("Travel Requests").items.getById(itemId).update(data)
      // .then(() => {
      //     resolve(itemId);
      // })
      // .catch(e => {
      //     reject(e);
      // });

      resolve(itemId);
    });
  }

  public async AddAttachments(formState: IAddAttachmentState, formKey: string) {
    return new Promise<any>(async (resolve, reject) => {
      let web = new Web(this.pc.web.absoluteUrl);
      if (formState.files.length > 0) {
        for (let i = 0; i < formState.files.length; i++) {
          try {
            let unique = getRandomString(4);
            let file = formState.files[i];
            let ext = file.name.substring(file.name.lastIndexOf("."));
            let tmpFileName = file.name.substring(
              0,
              file.name.lastIndexOf(".")
            );
            let fName = tmpFileName + "_" + unique + ext;
            let result = await web
              .getFolderByServerRelativeUrl("FormAttachments")
              .files.add(fName, file, true);
            let item = await result.file.getItem();
            await item.update({
              FormKey: formKey,
            });
            resolve("success");
          } catch (e) {
            console.log(
              "Error loading attachment or setting associted metadata"
            );
            reject(e);
          }
        }
      }
    });
  }

  public async GetAttachments(formKey: string) {
    return new Promise<any>(async (resolve, reject) => {
      let web = new Web(this.pc.web.absoluteUrl);
      web.lists
        .getByTitle("Form Attachments")
        .items.select("FileLeafRef", "Id")
        .filter("FormKey eq '" + formKey + "'")
        .top(500)
        .get()
        .then(async (res) => {
          resolve(res);
        })
        .catch((e) => {
          console.log("Error getting target items");
          reject(e);
        });
    });
  }

  public async RemoveAttachment(attId: string) {
    return new Promise<any>(async (resolve, reject) => {
      let web = new Web(this.pc.web.absoluteUrl);
      await web.lists
        .getByTitle("Form Attachments")
        .items.getById(Number(attId))
        .delete();
      resolve("success");
    });
  }

  public async GetApprovers(employee: string) {
    return new Promise<any>(async (resolve, reject) => {
      let web = new Web(this.pc.web.absoluteUrl);
      web.lists
        .getByTitle("Travel Requests - Approval Matrix")
        .items.expand(
          "Employee/Name",
          "Admin",
          "SectionHead",
          "Secretary",
          //"Undersecretary",
          "DeputyUndersecretary",
          "Budget",
          "AcctMgr1",
          "AcctMgr2",
          "AcctAdmin"
        )
        .select(
          "Employee/Name",
          "Admin/Name",
          "Admin/Id",
          "Admin/Title",
          "Admin/UserName",
          "AcctAdmin/UserName",
          "AcctAdmin/Id",
          "AcctAdmin/Title",
          "SectionHead/UserName",
          "SectionHead/Id",
          "SectionHead/Title",
          "Secretary/UserName",
          //"Undersecretary/UserName",
          //"Undersecretary/Id",
          "DeputyUndersecretary/UserName",
          "DeputyUndersecretary/Id",
          "Secretary/Title",
          "Secretary/Id",
          //"Undersecretary/Title",
          "DeputyUndersecretary/Title",
          "Budget/UserName",
          "Budget/Title",
          "Budget/Id",
          "AcctMgr1/UserName",
          "AcctMgr2/UserName",
          "Agency",
          "PersonnelNo"
        )
        .filter("Employee/Name eq '" + encodeURIComponent(employee) + "'")
        .top(500)
        .get()
        .then(async (res) => {
          if (res.length > 0) {
            resolve(res[0]);
          }
        })
        .catch((e) => {
          console.log("Error getting target items");
          reject(e);
        });
    });
  }

  ///_api/web/lists/GetByTitle('customlisttitle')/items?$select=ID,Owner/EMail,Owner/FirstName,Owner/LastName&$filter=Owner/LoginName eq "+encodeURI('i:0#.f|membership|spintboxtest@fluidigm.com')+"&$expand=Owner/LoginName

  //end data service
}
