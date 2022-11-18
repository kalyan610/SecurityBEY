import * as React from 'react';
//import styles from './SecurityBey.module.scss';
import styles from './SecurityBey.module.scss';
import { getGUID } from "@pnp/common";
import { ISecurityBeyProps } from './ISecurityBeyProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Service from './Service';
import { ChoiceGroup, IChoiceGroupOption, textAreaProperties, Stack, IStackTokens, StackItem, IStackStyles, TextField, CheckboxVisibility, BaseButton, Button } from 'office-ui-fabric-react';
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { Checkbox, PrimaryButton } from 'office-ui-fabric-react';
//import {Checkbox, Label, PrimaryButton} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { stringIsNullOrEmpty, isArray, objectDefinedNotNull } from '@pnp/common/util';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Items } from 'sp-pnp-js';

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 200 },
};
const sectionStackTokens: IStackTokens = { childrenGap: 10 };
const stackTokens = { childrenGap: 80 };
const stackStyles: Partial<IStackStyles> = { root: { padding: 10 } };
const stackButtonStyles: Partial<IStackStyles> = { root: { Width: 20 } };
var arr = [];
let adminexists = 'false';
let RecordId = '';
let drplocval: any;
let itemId = '';
let appredate = null;
let mylookupLocation = null;
let myproject = null;
let RedirectUrl = '';
let ApproverEmail = null;
let Aproverlogin = null;

export interface IEditFormProps {
}

export interface IEditFormState {
  layoutOption: string;
  list: any;
  flag: boolean;
  TypedEnterflag: boolean;
  TotalPages: number;
  myRecIndex: number;
}

interface IPeoplePickerControlInfo {
  Title: string;
}

interface ILookupFieldInfo {
  Id: number;
  Title: string;
}

interface IUserRecords {
  AccessCard: string;
  ApproverName: string;
  ApproverSignId: string;
  ApproversEmployeeCode: string;

  CapcoEmployeeCode: string;
  CLEARANCELIST: string[];
  ContractorID: string;
  //dtdoj: Date;
  //RequestDate: string;
  AdminAPProverName: IPeoplePickerControlInfo;
  RequestorName: IPeoplePickerControlInfo;
  Requestorsign: IPeoplePickerControlInfo;
  ApprovalSignature: IPeoplePickerControlInfo;
  Title: string;
  Location: ILookupFieldInfo;
  dtofbiometric: any,
  ApprovalSignatureDate: any,
  AdminComments: string

}

export interface ISecurityBeyState {
  isLoading: boolean;
  userRecords: IUserRecords;
  Status: any;
  EmpName: any;
  dtdoj: any;
  EmpID: any;
  ItemInfo: any;
  allchecklistitems: any;
  flag: boolean;
  ApprovalSignature: any;
  ApprovalSignatureDate: any;
  bloodgroup: any;
  emconnum: any;
  Singleuser: boolean;
  CapcoEmployeeCode: any;
  ContractorID: any;
  ProjectandLocation: any;
  Requestorsign: any;
  RequestorName: any;
  AdminAPProverName: any;
  ApproverName: any;
  ApproversEmployeeCode: any;
  AdminSignPeople: any;
  dtofbiometric: any;
  AdminExsists: boolean;
  AdminComments: string;
  multiValueCheckbox: any;
  choiceValues: string[];
  TotalcleraancelIst: any;
  QueryStringSIDValue: string;
  projectLocation: IDropdownOption;
  ApproverDated: any;
  LocationListItems: IDropdownOption[];
}

export default class SecurityBey extends React.Component<ISecurityBeyProps, ISecurityBeyState> {
  public _service: Service;
  protected ppl: PeoplePicker;
  public constructor(props: ISecurityBeyProps) {
    super(props);
    this.state = {
      isLoading: true,
      userRecords: null,
      Status: '',
      ItemInfo: "",
      allchecklistitems: 'BW Bay (Gr Floor)',
      flag: false,
      EmpName: null,
      dtdoj: null,
      EmpID: null,
      bloodgroup: null,
      emconnum: null,
      CapcoEmployeeCode: '',
      ContractorID: '',
      ProjectandLocation: '',
      Singleuser: false,
      ApproversEmployeeCode: null,
      ApproverName: null,
      AdminExsists: false,
      dtofbiometric: null,
      TotalcleraancelIst: [],
      RequestorName: [],
      AdminAPProverName: [],
      Requestorsign: [],
      ApprovalSignature: [],
      ApprovalSignatureDate: null,
      AdminSignPeople: [],
      AdminComments: null,
      multiValueCheckbox: [],
      choiceValues: [],
      QueryStringSIDValue: this.getParam('SID'),
      projectLocation: null,
      ApproverDated: null,
      LocationListItems: null,

    };

    this._service = new Service(this.props.url, this.props.context);


    RedirectUrl = "https://capcoinc.sharepoint.com/sites/INBAFacility/"
  }

  public render(): React.ReactElement<ISecurityBeyProps> {
    return (
      <React.Fragment>
        {this.state.isLoading
          ? <div>Loading....</div>
          : <Stack>
            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                <b>{'Requestor (Last Name, First, MI)'}</b>
              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <StackItem >
                <b>{'Capco Employee Code'}</b>
              </StackItem>
            </Stack><br></br>
            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                <PeoplePicker
                  context={this.props.context as any}
                  //titleText="User Name"
                  personSelectionLimit={1}
                  showtooltip={true}
                  required={true}
                  disabled={this.state.AdminExsists == true ? true : false}
                  onChange={this._getPeoplePickerItems2.bind(this)}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  defaultSelectedUsers={this.state.userRecords.RequestorName ? [this.state.userRecords.RequestorName.Title] : []}
                  ref={c => (this.ppl = c)}
                  resolveDelay={1000} />




              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <StackItem >
                <TextField
                  id="CapcoEmployeeCode"
                  name="CapcoEmployeeCode"
                  autoAdjustHeight
                  defaultValue={this.state.userRecords.CapcoEmployeeCode}
                  disabled={this.state.AdminExsists == true ? true : false}
                  // onChange={(e, newValue: string) => { this.setState({ CapcoEmployeeCode: newValue }); }}
                  onChange={(e, newValue: string) => { this.setState({ userRecords: { ...this.state.userRecords, CapcoEmployeeCode: newValue } }) }}
                />


              </StackItem>
            </Stack><br></br>
            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                <b>Contractor ID # (If Any):</b>
              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <StackItem >
                <b>Access Card # (Internal Use Only)</b>
              </StackItem>
            </Stack><br></br>

            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                <TextField
                  id="ContractorID"
                  name="ContractorID"
                  autoAdjustHeight
                  defaultValue={this.state.userRecords.ContractorID}
                  disabled={this.state.AdminExsists == true ? true : false}
                  // onChange={(e, newValue: string) => { this.setState({ ContractorID: newValue }); }}
                  onChange={(e, newValue: string) => { this.setState({ userRecords: { ...this.state.userRecords, ContractorID: newValue } }) }}
                />
              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <StackItem >

                <TextField
                  id="AccessCard"
                  name="AccessCard"
                  autoAdjustHeight
                  defaultValue={this.state.userRecords.AccessCard}
                  disabled={this.state.AdminExsists == true ? true : false}
                  onChange={(e, newValue: string) => { this.setState({ userRecords: { ...this.state.userRecords, AccessCard: newValue } }) }}
                />
              </StackItem>
            </Stack><br></br>
            <Stack>
              <b>Declaration:<br></br>
                I understand that access to the Capco information resources requested is being granted for purpose of conducting business
                associated with Capco and/or its’ customers. I am aware that I am responsible to ensure that this access is in compliance with
                all provisions of Capco Security and Procedures. Failure to comply with these policies and procedures could result in
                disciplinary action, criminal prosecution, and termination of any contractual or employment relationship with Capco.</b>
            </Stack><br></br>
            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                <b>Requestor’s Signature:</b>
              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <StackItem >
                <b>Date:</b>
              </StackItem>
            </Stack><br></br>

            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                <PeoplePicker
                  context={this.props.context as any}
                  //titleText="User Name"
                  personSelectionLimit={1}
                  showtooltip={true}
                  required={true}
                  disabled={this.state.AdminExsists == true ? true : false}
                  onChange={this._getPeoplePickerItems3}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  defaultSelectedUsers={this.state.userRecords.Requestorsign ? [this.state.userRecords.Requestorsign.Title] : []}
                  ref={c => (this.ppl = c)}
                  resolveDelay={1000} />

              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <StackItem >
                <DateTimePicker
                  dateConvention={DateConvention.Date}
                  disabled={this.state.AdminExsists == true ? true : false}
                  showLabels={false}
                  value={this.state.userRecords.dtofbiometric}
                  onChange={(date: any) => { this.setState({ userRecords: { ...this.state.userRecords, dtofbiometric: date } }) }} />


              </StackItem>
            </Stack><br></br>
            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                <b>Project and Location</b>
              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <StackItem >
                <b> Approver (Last Name, First, MI)</b>
              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                  <StackItem >
                <b>Approver’s Employee Code</b>
              </StackItem>
            </Stack><br></br>

            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                <Dropdown
                  placeholder={"Select Location"}
                  options={this.state.LocationListItems}
                  multiSelect={false}
                  styles={dropdownStyles}
                  defaultSelectedKey={objectDefinedNotNull(this.state.userRecords.Location) ? this.state.userRecords.Location.Id : 0}
                  //selectedKey={objectDefinedNotNull(this.state.userRecords.Location)?this.state.userRecords.Location.Id:0}

                  onChange={this.handleprojecteType} disabled={this.state.AdminExsists == true ? true : false} />
                <br /><br></br>

              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  {/* <StackItem className={styles.commonstyle}>
                <TextField
                  id="ApproverName"
                  name="ApproverName"
                  autoAdjustHeight
                  value={this.state.userRecords.ApproverName}
                  disabled={true}
                  onChange={(e, newValue: string) => { this.setState({ userRecords: { ...this.state.userRecords, ApproverName: newValue } }) }}
                />
              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; */}

              <StackItem >
                <PeoplePicker
                  context={this.props.context as any}
                  //titleText="User Name"
                  personSelectionLimit={10}
                  showtooltip={true}
                  required={true}
                  disabled={true}
                  //disabled={this.state.AdminExsists == true ? true : false}
                  onChange={this._getPeoplePickerItems5}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}

                  defaultSelectedUsers={this.state.userRecords.AdminAPProverName ? [this.state.userRecords.AdminAPProverName.Title] : []}


                  ref={c => (this.ppl = c)}
                  resolveDelay={1000} />

              </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <StackItem >
                <TextField
                  id="ApproversEmployeeCode"
                  name="ApproversEmployeeCode"
                  autoAdjustHeight
                  value={this.state.userRecords.ApproversEmployeeCode}
                  disabled={true}
                  onChange={(e, newValue: string) => { this.setState({ userRecords: { ...this.state.userRecords, ApproversEmployeeCode: newValue } }) }}
                />
              </StackItem>
            </Stack><br></br>

            <Stack horizontal tokens={sectionStackTokens}>
              <b>CLEARANCE LIST (Check all clearances to be added):</b><br></br>
              <StackItem >
              </StackItem>
            </Stack><br></br>
            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                {this.state.choiceValues.map((choiceValue: string) => {
                  const isItemAvailable: boolean = this.state.userRecords.CLEARANCELIST.some(s => s == choiceValue);
                  return (
                    <div style={{ margin: "2px", padding: "3px" }}>
                      <Checkbox style={{ margin: "2px", padding: "3px" }} defaultChecked={isItemAvailable} label={choiceValue} onChange={this._onCheckboxMultiChecked} disabled={this.state.AdminExsists == true ? true : false} />
                    </div>
                  );
                }
                )}
              </StackItem>
            </Stack><br></br>
            <Stack horizontal tokens={sectionStackTokens}>
              <StackItem >
                <PrimaryButton onClick={this.Save} disabled={this.state.AdminExsists == true ? true : false}  >{'Submit'}</PrimaryButton>
              </StackItem>
            </Stack><br></br>
            {this.state.AdminExsists == true &&
              <Stack>
                <Stack horizontal tokens={sectionStackTokens}>
                  <StackItem >
                    <b>Approval Signature:</b>
                  </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <StackItem >
                    <b>Date:</b>
                  </StackItem>
                </Stack><br></br>

                <Stack horizontal tokens={sectionStackTokens}>
                  <StackItem >
                    <PeoplePicker
                      context={this.props.context as any}
                      //titleText="User Name"
                      personSelectionLimit={1}
                      showtooltip={true}
                      required={true}
                      onChange={this._getPeoplePickerItems7}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      defaultSelectedUsers={this.state.userRecords.ApprovalSignature ? [this.state.userRecords.ApprovalSignature.Title] : []}
                      ref={c => (this.ppl = c)}
                      resolveDelay={1000} />
                  </StackItem>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <StackItem >
                    <DateTimePicker
                      dateConvention={DateConvention.Date}
                      showLabels={false}
                      value={this.state.userRecords.ApprovalSignatureDate}
                      onChange={(date: any) => { this.setState({ userRecords: { ...this.state.userRecords, ApprovalSignatureDate: date } }) }} />
                    {/* // value={this.state.ApproverDated}
                      //onChange={this.handlechangedtrequestordate} />  */}
                  </StackItem><br></br>
                </Stack><br></br>

                <Stack horizontal tokens={sectionStackTokens}>
                  <StackItem >
                    <b>Admin Comments</b>
                  </StackItem>
                </Stack><br></br>
                
                  <div className={styles.textAreacss}>
                    <textarea id="txtAdminComments" className={styles.textAreacss}

                      defaultValue={this.state.userRecords.AdminComments}
                      onChange={this.changeAdminComments.bind(this)}
                    //onChange={(e:[], newValue: string) => { this.setState({ userRecords: { ...this.state.userRecords, AccessCard: newValue } }) }}
                    ></textarea>
                    </div>
                  
                

                <Stack horizontal tokens={sectionStackTokens}>
                  <StackItem >
                    <PrimaryButton onClick={e => this.Approved(e)}>Approve</PrimaryButton>&nbsp;&nbsp;&nbsp;&nbsp;
                <PrimaryButton onClick={e => this.Rejected(e)}>Reject</PrimaryButton>
                  </StackItem><br>
                  </br>


                </Stack><br></br>

              </Stack>

            }
          </Stack>
        }
      </React.Fragment>
    );
  }
  public async componentDidMount() {
    try {


      const choiceValues: string[] = await this.getCheckboxChoices("CLEARANCELIST");
      const userRecords: IUserRecords = await this.getHRandAdminGroupUserorNot();
      const ddlLocationItems: IDropdownOption[] = await this._getProjectandLocation();
      this.setState({ isLoading: false, choiceValues, userRecords, LocationListItems: ddlLocationItems });
      await this.Approved;
    } catch (error) {
      console.error(error);
    }
  }


  public getParam(name: string) {
    name = name.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]");
    var regexS = "[\\?&]" + name + "=([^&#]*)";
    var regex = new RegExp(regexS);
    var results = regex.exec(window.location.href);
    if (results == null)
      return "";
    else
      return results[1];
  }



  private async _getPeoplePickerItems2(items: any[]) {
    console.log('Items:', items);
    if (items.length > 0) {

      let userInfo = this._service.getUserByLogin(items[0].loginName).then((info) => {
        this.setState({ RequestorName: info });
        console.log(info);
      });
    }
    else {
      this.setState({ RequestorName: null });
    }
  }

  private changeAdminComments(data: any): void {

    this.setState({ AdminComments: data.target.value });
  }


  private async _getPeoplePickerItems4(items: any[]) {
    console.log('Items:', items);
    if (items.length > 0) {

      let userInfo = this._service.getUserByLogin(items[0].loginName).then((info) => {
        this.setState({ RequestorName: info });
        console.log(info);
      });
    }
    else {
      this.setState({ RequestorName: null });
    }
  }

  // private _getPeoplePickerItems3 = async (items: any[]): Promise<void> => {
  //   return new Promise<void>((resolve, reject) => {
  //     try {
  //       if (items.length > 0) {
  //         this._service.getUserByLogin(items[0].loginName)
  //           .then((info) => {
  //             this.setState({ userRecords: { ...this.state.userRecords, Requestorsign: info } })
  //           });
  //       }
  //       else {
  //         this.setState({ Requestorsign: null });
  //       }
  //       //this.ppl.onChange([]);
  //     } catch (error) {

  //     }
  //   })
  // }

  private _getPeoplePickerItems3 = async (items: any[]): Promise<void> => {
    console.log('Items:', items);
    if (items.length > 0) {

      let userInfo = this._service.getUserByLogin(items[0].loginName).then((info) => {
        this.setState({ Requestorsign: info });
        console.log(info);
      });
    }
    else {
      this.setState({ Requestorsign: null });
    }
  }

  private _getPeoplePickerItems5 = async (items: any[]): Promise<void> => {
    console.log('Items:', items);
    if (items.length > 0) {

      let userInfo = this._service.getUserByLogin(items[0].loginName).then((info) => {
        this.setState({ AdminAPProverName: info });
        console.log(info);
      });
    }
    else {
      this.setState({ AdminAPProverName: null });
    }
  }
  private _getPeoplePickerItems7 = async (items: any[]): Promise<void> => {
    console.log('Items:', items);
    if (items.length > 0) {

      let userInfo = this._service.getUserByLogin(items[0].loginName).then((info) => {
        this.setState({ ApprovalSignature: info });
        console.log(info);
      });
    }
    else {
      this.setState({ ApprovalSignature: null });
    }
  }


  private async _getApproverName(items: any[]) {
    console.log('Items:', items);
    if (items.length > 0) {

      let userInfo = this._service.getUserByLogin(items[0].loginName).then((info) => {
        this.setState({ AdminSignPeople: info });
        console.log(info);
      });
    }
    else {
      this.setState({ AdminSignPeople: null });
    }
    //this.ppl.onChange([]);

  }
  // private async _getApproversign(items: any[]) {
  //   console.log('Items:', items);
  //   if (items.length > 0) {

  //     let userInfo = this._service.getUserByLogin(items[0].loginName).then((info) => {
  //       this.setState({ ApprovalSignature: info });
  //       console.log(info);
  //     });
  //   }
  //   else {
  //     this.setState({ ApprovalSignature: null });
  //   }
  //   //this.ppl.onChange([]);

  // }

  private handleprojecteType = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): Promise<any> => {
    try {
      const ItemInfo: { ApproverEMPCode: string, Approver: IPeoplePickerControlInfo } = await this._getItemByTitle(item.text);
      this.setState({ projectLocation: item, userRecords: { ...this.state.userRecords, ApproversEmployeeCode: ItemInfo.ApproverEMPCode, AdminAPProverName: ItemInfo.Approver } });
    } catch (error) {
      console.error(error);
    }
  }

  private _getItemByTitle = (ItemTitle: any): Promise<{ ApproverEMPCode: string, Approver: IPeoplePickerControlInfo }> => {
    return new Promise<{ ApproverEMPCode: string, Approver: IPeoplePickerControlInfo }>(async (resolve, reject) => {
      try {
        const selectedList = 'ProjectandLocation';
        const Item: any[] = await sp.web.lists.getByTitle(selectedList).items.select("ApproverEMPCode,Approver/EMail").expand("Approver").top(1).filter("Title eq '" + ItemTitle + "'").get();
        const projectLocationItem: any = Item[0];
        resolve({
          ApproverEMPCode: projectLocationItem['ApproverEMPCode'] ?? '',
          Approver: (projectLocationItem['Approver'] != null) ? await this._getPeoplePickerItems22(projectLocationItem.Approver.EMail) : null,



        });
      } catch (error) {
        console.log(error);
        reject(error);
      }
    })
  }

  private _getProjectandLocation = (): Promise<IDropdownOption[]> => {
    return new Promise<IDropdownOption[]>(async (sucess, failure) => {
      try {
        const projectLocations: any[] = await this._service.GetProjectandLocation();
        const ddlLocationItems: IDropdownOption[] = projectLocations.map<IDropdownOption>((ploc: any) => ({ key: ploc.ID, text: ploc.Title }));
        sucess(ddlLocationItems);
      } catch (error) {
        console.error(error);
        failure(error);
      }
    });
  }

  public handlechangedtrequestordate = (date: any) => {
    this.setState({ dtofbiometric: date });
    this.setState({ ApprovalSignatureDate: date })
    this.setState({ ApproverDated: date });
    this.setState({ ApprovalSignatureDate: date })
  }


  private handleChangeApproverName(data: any): void {

    this.setState({ ApproverName: data.target.value });
  }
  private handleChangeApproversEmployeeCode(data: any): void {


    this.setState({ ApproversEmployeeCode: data.target.value });
  }




  private _onCheckboxMultiChecked = async (ev: React.FormEvent<HTMLElement>, isChecked: boolean): Promise<void> => {
    try {
      let updtedSelectedChoiceValues: string[] = [...this.state.userRecords.CLEARANCELIST];

      const choiceValue: string = ev.currentTarget["ariaLabel"];
      if (isChecked) {
        updtedSelectedChoiceValues.push(choiceValue);
      } else {
        updtedSelectedChoiceValues = this.state.userRecords.CLEARANCELIST.filter(sc => sc !== choiceValue);
      }

      this.setState(prev => ({
        userRecords: {
          ...prev.userRecords,
          CLEARANCELIST: updtedSelectedChoiceValues
        }
      }));
    } catch (error) {
      console.error(error);
    }
  }


  public getCheckboxChoices = (fieldname: string): Promise<string[]> => {
    return new Promise<string[]>((resolve, reject) => {
      try {
        sp.web.lists.getByTitle("SecuredBayAccess").fields
          .getByInternalNameOrTitle(fieldname)
          .select('Choices')
          .get()
          .then((field) => {
            const multivaluechoices: string[] = [];
            for (var i = 0; i < field["Choices"].length; i++) {
              multivaluechoices.push(field["Choices"][i]?.toString());
            }
            resolve(multivaluechoices);
          });
      } catch (error) {
        console.error(error);
        reject(error);
      }
    })
  }
  private getHRandAdminGroupUserorNot = async (): Promise<IUserRecords> => {


    let userRecords: IUserRecords = null;

    const loginuser = await this._service.getCurrentUser();
    let mylogin = [];
    mylogin = loginuser.Email;
    userRecords = await this.getuserrecords();
    ApproverEmail = Aproverlogin;

    const mycurgroups: any[] = await this._service.getCurrentUserSiteGroups();
    const isCurrentUserAdmin: boolean = mycurgroups.some((grp: any) => grp.Title === 'Secured Bay Access Approvers');

    if (mylogin == ApproverEmail || isCurrentUserAdmin) {
      this.setState({ AdminExsists: true });

    }
    else {

      this.setState({ AdminExsists: false });
    }







    return userRecords;
  }

  private getuserrecords = async (): Promise<IUserRecords> => {
    const defaultVaues: IUserRecords = {
      AccessCard: '',
      CapcoEmployeeCode: '',
      ContractorID: '',
      //dtdoj: new Date,
      RequestorName: { Title: '' },
      Requestorsign: { Title: '' },
      ApprovalSignature: { Title: '' },
      ApproverName: '',
      ApproversEmployeeCode: '',
      ApproverSignId: '',
      CLEARANCELIST: [],
      AdminAPProverName: { Title: '' },
      AdminComments: '',

      Title: '',
      // RequestDate: '',
      Location: null,
      dtofbiometric: null,
      ApprovalSignatureDate: null
    };

    if (stringIsNullOrEmpty(this.state.QueryStringSIDValue)) {
      return defaultVaues;
    } else {
      const ItemInfo = await this._service.getItemByID(this.state.QueryStringSIDValue);
      if (ItemInfo.Title != '') {


        Aproverlogin = ItemInfo.ApproverNames.EMail;
        mylookupLocation = ItemInfo.Location.Id;
        let strdoj = ItemInfo.RequestDate.split('T');
        strdoj[0].replace("-", "/");
        let mainstr = strdoj[0].replace("-", "/");
        let strToDate = new Date(mainstr);
        if (ItemInfo.ApproverDate != null) {
          let Approerdate = ItemInfo.ApproverDate?.split('T');
          Approerdate[0].replace("-", "/");
          let mainstr1 = Approerdate[0].replace("-", "/");
          appredate = new Date(mainstr1);
          // this.setState({:strToDate})
        }
        return ({
          AccessCard: ItemInfo['AccessCard'] ?? '',
          CapcoEmployeeCode: ItemInfo['CapcoEmployeeCode'] ?? '',
          ContractorID: ItemInfo['ContractorID'] ?? '',
          //dtdoj: new Date(mainstr),
          RequestorName: (ItemInfo.RequestorNameId != null) ? await this._getPeoplePickerItems20(ItemInfo.RequestorName.EMail) : null,
          Requestorsign: (ItemInfo.EMPSIGNId != null) ? await this._getPeoplePickerItems20(ItemInfo.EMPSIGN.EMail) : null,
          ApproverName: ItemInfo['ApproverName'] ?? '',
          ApproversEmployeeCode: ItemInfo['ApproverEMPCode'] ?? '',
          //dtdoj:ItemInfo['RequestDate']??'',

          ApproverSignId: ItemInfo['ApproverSign'] ?? '',
          CLEARANCELIST: ItemInfo['CLEARANCELIST'] ?? [],
          Title: ItemInfo['Title'] ?? '',
          dtofbiometric: strToDate ?? '',
          ApprovalSignatureDate: appredate ?? null,
          Location: ItemInfo["Location"] ?? { Id: 0, Title: '' },
          ApprovalSignature: (ItemInfo.ApproverSignId != null) ? await this._getPeoplePickerItems21(ItemInfo.ApproverSign.EMail) : null,
          AdminAPProverName: (ItemInfo.ApproverNamesId != null) ? await this._getPeoplePickerItems22(ItemInfo.ApproverNames.EMail) : null,
          AdminComments: ItemInfo['AdminComments'] ?? ''
        });
      } else {
        return defaultVaues;
      }
    }
  }
  private _getPeoplePickerItems20 = async (UserEmail: string): Promise<IPeoplePickerControlInfo> => {
    return new Promise<IPeoplePickerControlInfo>((success, failure) => {
      try {
        if (UserEmail.length > 0) {
          this._service.getUserByEmail(UserEmail)
            .then((info) => {
              success({ Title: info.Title })
              this.setState({ RequestorName: info, Requestorsign: info });
            });
        } else {
          success({ Title: '' });
        }
      } catch (error) {
        console.error(error);
        failure(error);
      }
    })
  }

  private _getPeoplePickerItems21 = async (UserEmail: string): Promise<IPeoplePickerControlInfo> => {
    return new Promise<IPeoplePickerControlInfo>((success, failure) => {
      try {
        if (UserEmail.length > 0) {
          this._service.getUserByEmail(UserEmail)
            .then((info) => {
              success({ Title: info.Title })
              this.setState({ ApprovalSignature: info });
            });
        } else {
          success({ Title: '' });
        }
      } catch (error) {
        console.error(error);
        failure(error);
      }
    })
  }

  private _getPeoplePickerItems22 = async (UserEmail: string): Promise<IPeoplePickerControlInfo> => {
    return new Promise<IPeoplePickerControlInfo>((success, failure) => {
      try {
        if (UserEmail.length > 0) {
          this._service.getUserByEmail(UserEmail)
            .then((info) => {
              success({ Title: info.Title })
              this.setState({ AdminAPProverName: info });
            });
        } else {
          success({ Title: '' });
        }
      } catch (error) {
        console.error(error);
        failure(error);
      }
    })
  }

  //   private Approved=(e: React.MouseEvent<HTMLAnchorElement | HTMLButtonElement | HTMLDivElement | BaseButton | Button | HTMLSpanElement, MouseEvent>)  => {
  //     window.location.href = "https://google.com/contact";

  // }

  private Approved = (e: React.MouseEvent<HTMLAnchorElement | HTMLButtonElement | HTMLDivElement | BaseButton | Button | HTMLSpanElement, MouseEvent>): Promise<void> => {
    return new Promise<void>(async (success, failure) => {
      try {

        e.preventDefault();
        if (this.state.ApprovalSignature == null || this.state.ApprovalSignature == "") {
          alert('Please enter Approval Signature');
        } else if (this.state.userRecords.ApprovalSignatureDate == null || this.state.userRecords.ApprovalSignatureDate == "") {
          alert('Please enter Date');
        }
        let ApprDate = this.state.userRecords.ApprovalSignatureDate.getDate() + 1;
        let month4 = (this.state.userRecords.ApprovalSignatureDate.getMonth() + 1);
        let year4 = (this.state.userRecords.dtofbiometric.getFullYear());
        let FinalDateofresign = month4 + '/' + this.state.userRecords.ApprovalSignatureDate.getDate() + '/' + year4;
        let status = "Approved";
        let ApproversignEmailid = (this.state.ApprovalSignature == null ? 0 : this.state.ApprovalSignature.Id);


        const { CapcoEmployeeCode, QueryStringSIDValue } = this.state;





        await sp.web.lists.getByTitle("SecuredBayAccess").items
          .getById(+QueryStringSIDValue)
          .update({
            ApproverSignId: ApproversignEmailid,
            ApproverStatus: status,
            ApproverDate: FinalDateofresign,
            AdminComments: this.state.AdminComments

          }).then(i => {
            alert("Approved Successfully");
            success();
            window.location.href =RedirectUrl;
          }).catch((error) => { console.error(error); failure(error); })
      } catch (error) {
        console.error(error);
        failure(error);
      }
    });
  }



  private Rejected = (e: React.MouseEvent<HTMLAnchorElement | HTMLButtonElement | HTMLDivElement | BaseButton | Button | HTMLSpanElement, MouseEvent>): Promise<void> => {
    return new Promise<void>(async (success, failure) => {
      try {

        e.preventDefault();

        if (this.state.ApprovalSignature == null || this.state.ApprovalSignature == "") {
          alert('Please enter Approval Signature ');
        } else if (this.state.userRecords.ApprovalSignatureDate == null || this.state.userRecords.ApprovalSignatureDate == "") {
          alert('Please enter Date');
        }
        let ApprDate = this.state.userRecords.ApprovalSignatureDate.getDate() + 1;
        let month4 = (this.state.userRecords.ApprovalSignatureDate.getMonth() + 1);
        let year4 = (this.state.userRecords.dtofbiometric.getFullYear());
        let FinalDateofresign = month4 + '/' + this.state.userRecords.ApprovalSignatureDate.getDate() + '/' + year4;
        let status = "Rejected";
        let ApproversignEmailid = (this.state.ApprovalSignature == null ? 0 : this.state.ApprovalSignature.Id);


        const { CapcoEmployeeCode, QueryStringSIDValue } = this.state;






        await sp.web.lists.getByTitle("SecuredBayAccess").items
          .getById(+QueryStringSIDValue)
          .update({
            ApproverSignId: ApproversignEmailid,
            ApproverStatus: status,
            ApproverDate: FinalDateofresign,
            AdminComments: this.state.AdminComments

          }).then(i => {
            alert("Rejected Successfully");
            
           success();
            window.location.href =RedirectUrl;
          }).catch((error) => { console.error(error); failure(error); })

      } catch (error) {
        console.error(error);
        failure(error);
      }
    });
  }



  private Save = (e: React.MouseEvent<HTMLAnchorElement | HTMLButtonElement | HTMLDivElement | BaseButton | Button | HTMLSpanElement, MouseEvent>): Promise<void> => {
    return new Promise<void>(async (success, failure) => {
      try {
        e.preventDefault();

        //destructuring
        const { CapcoEmployeeCode, QueryStringSIDValue } = this.state;
        if (stringIsNullOrEmpty(QueryStringSIDValue)) {
          if (this.state.RequestorName == null || this.state.RequestorName == "") {
            alert('Please enter Requestor Name');
          } else if (this.state.userRecords.CapcoEmployeeCode == null || this.state.userRecords.CapcoEmployeeCode == "") {
            alert('Please enter  Capco EMP ID');
          } //else if (this.state.userRecords.ContractorID == null || this.state.userRecords.ContractorID == "") {
            //alert('Please enter  Contractor ID');} 
          else if (this.state.userRecords.AccessCard == null || this.state.userRecords.AccessCard == "") {
            alert('Please enter  Access Card');
          } else if (this.state.Requestorsign == null || this.state.Requestorsign == "") {
            alert('Please enter  Requestor’s Signature');
          } else if (this.state.userRecords.dtofbiometric == null || this.state.userRecords.dtofbiometric == "") {
            alert('Please Select  Date');
          } else if (!objectDefinedNotNull(this.state.projectLocation)) {
            alert('Please Select Project and Location');
          } else if (isArray(this.state.userRecords.CLEARANCELIST) && this.state.userRecords.CLEARANCELIST.length == 0) {
            alert('Please Select  CLEARANCE LIST');
          } else {
            //add item Logic
            let myproject = this.state.projectLocation;
            let status = "Pending";
            let RequestorEmailid = (this.state.RequestorName == null ? 0 : this.state.RequestorName.Id);
            let requestoremailsignig = (this.state.Requestorsign == null ? 0 : this.state.Requestorsign.Id);
            let ApprovernewEmail = (this.state.AdminAPProverName == null ? 0 : this.state.AdminAPProverName.Id);
            let date4 = this.state.userRecords.dtofbiometric.getDate() + 1;
            let month4 = (this.state.userRecords.dtofbiometric.getMonth() + 1);
            let year4 = (this.state.userRecords.dtofbiometric.getFullYear());
            let FinalDateofresign = month4 + '/' + this.state.userRecords.dtofbiometric.getDate() + '/' + year4;

            await sp.web.lists.getByTitle("SecuredBayAccess").items
              .add({
                Title: getGUID(),
                RequestorNameId: RequestorEmailid,
                ApproverNamesId: ApprovernewEmail,
                EMPSIGNId: requestoremailsignig,
                CapcoEmployeeCode: this.state.userRecords.CapcoEmployeeCode,
                ContractorID: this.state.userRecords.ContractorID,
                AccessCard: this.state.userRecords.AccessCard,
                //RequestDate: this.state.userRecords.dtdoj,
                RequestDate: FinalDateofresign,
                CLEARANCELIST: { results: this.state.userRecords.CLEARANCELIST },
                LocationId: this.state.projectLocation.key,
                //LocationId: this.state.userRecords.Location.Id,
                ApproverName: this.state.userRecords.ApproverName,
                ApproverEMPCode: this.state.userRecords.ApproversEmployeeCode,
                ApproverStatus: status
              }).then(i => {
                alert("Submitted Successfully");
               
                success();
                window.location.href =RedirectUrl;

              }).catch((error) => { console.error(error); failure(error); })
          }
        } else {
          // validations
          if (stringIsNullOrEmpty(this.state.userRecords.AccessCard)) {
            alert('Please enter  Access Card');
          }
          if (stringIsNullOrEmpty(this.state.RequestorName)) {
            alert('Please enter  Requestor Name');
          }
          if (stringIsNullOrEmpty(this.state.userRecords.CapcoEmployeeCode)) {
            alert('Please enter  Capco EMP ID');
          }
          if (stringIsNullOrEmpty(this.state.Requestorsign)) {
            alert('Please enter  Requestor Sign ');
          }


          if (isArray(this.state.userRecords.CLEARANCELIST) && this.state.userRecords.CLEARANCELIST.length == 0) {
            alert('Please enter  CLEARANCELIST');
          }
          else {
            //Update Logic
            if (this.state.projectLocation == null) {
              myproject = mylookupLocation;
            }
            else {
              myproject = this.state.projectLocation.key;

            }

            let status = "Pending";
            let RequestorEmailid = (this.state.RequestorName == null ? 0 : this.state.RequestorName.Id);
            let requestoremailsignig = (this.state.Requestorsign == null ? 0 : this.state.Requestorsign.Id);
            let date4 = this.state.userRecords.dtofbiometric.getDate() + 1;
            let month4 = (this.state.userRecords.dtofbiometric.getMonth() + 1);
            let year4 = (this.state.userRecords.dtofbiometric.getFullYear());
            let FinalDateofresign = month4 + '/' + this.state.userRecords.dtofbiometric.getDate() + '/' + year4;

            await sp.web.lists.getByTitle("SecuredBayAccess").items
              .getById(+QueryStringSIDValue)
              .update({
                RequestorNameId: RequestorEmailid,
                EMPSIGNId: requestoremailsignig,
                CapcoEmployeeCode: this.state.userRecords.CapcoEmployeeCode,
                ContractorID: this.state.userRecords.ContractorID,
                AccessCard: this.state.userRecords.AccessCard,
                //RequestDate: this.state.userRecords.dtdoj,
                RequestDate: FinalDateofresign,
                CLEARANCELIST: { results: this.state.userRecords.CLEARANCELIST },
                // LocationId: this.state.projectLocation.key,
                LocationId: myproject,
                ApproverName: this.state.userRecords.ApproverName,
                ApproverEMPCode: this.state.userRecords.ApproversEmployeeCode,
                ApproverStatus: status
              }).then(i => {
                alert("Updated Successfully");
                success();
                window.location.href =RedirectUrl;
                
              }).catch((error) => { console.error(error); failure(error); })
          }
        }
      } catch (error) {
        console.error(error);
        failure(error);
      }
    });
  }
}
