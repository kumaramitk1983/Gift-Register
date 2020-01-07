import * as React from 'react';
import { IGiftRegisterProps ,Status } from './IGiftRegisterProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Pivot, PivotItem, PivotLinkFormat, PivotLinkSize , IPivotItemProps } from 'office-ui-fabric-react/lib/Pivot';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { sp, Field, ItemAddResult, ItemUpdateResult, AttachmentFileAddResult, Items  } from "@pnp/sp";
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Button, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { AttachmentFile, AttachmentFiles, AttachmentFileInfo } from '@pnp/sp/src/attachmentfiles';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Dialog , DialogFooter , DialogType} from 'office-ui-fabric-react/lib/Dialog'
import { string } from 'prop-types';
import {Icon} from 'office-ui-fabric-react/lib/Icon';


const DayPickerStrings: IDatePickerStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],

  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker'
};

export default class GiftRegister extends React.Component<IGiftRegisterProps, any> {
  constructor(props: IGiftRegisterProps) {
    super(props);
    this.state = {
        _selectedKey: 0,
        _showSpinner : false,
        _approverMatrix : [],
        _showApprover : false,
        _currentUserId : null,
        _nonStandardApproverMatrix : [],
        _contractualStatusOptions : [],
        firstshowerror : true,
        _managerError : false,

        _errorCheck : {
          firstTab : {
            EmployeeDivision : true ,
            showerror : false,
          },
          secondTab : {
            Item_x0020_Value : true,
            ItemType : true,
            ItemGivenReceived : true,
            ItemIndividualCompany : true,
            ListIndividualCompany : true,
            ContractualStatus : true,
            Item_x0020_Description : true,
            showerror : false,
            tabValid : false,
          },
          approverTab : {
            comment : true,
            showerror : false
          }
        },
        

        //List Fields
        ID : -1,
        Requestor_x0020_Name : "",
        // Requestor_x0020_Phone : "",
        Requestor_x0020_Email : "",
        Employee_x0020_Title : "",
        ManagerName : "",
        RequestForSomeone : false,
        Item_x0020_Government : false,
        IsAmountReasonable : true,
        GiftDate : new Date(),
        ContractualStatus : [],
        ItemGivenReceived : '',
        EmployeeDivision : '',
        Item_x0020_Description : '',
        ItemType : '',
        ItemIndividualCompany : '',
        Item_x0020_Value : '',
        ListIndividualCompany : '',
        NonStandardApproval : false,
        Status : '',
        comment : '',

        //People Picker
        RequestorId : null,
        ManagerId : null,
        Approver1Id : null,
        Approver2Id : [],
        RequestedById : null,
        CurrentApproverId : [],
        ActionedById : [],

      //Attachment
        AttachmentFiles: [],

        //Validation Error
        showValidationError : false,
  
        
    };

    

    this._OnInit();
    sp.setup({
      spfxContext: this.props.context
    });
    this._tabClick = this._tabClick.bind(this);
    this._customRenderer = this._customRenderer.bind(this);
  }


  private _formatChoices = (res: Field): IDropdownOption[] => {
    let _choice: IDropdownOption[] = [];
    res["Choices"].forEach(element => {
      _choice.push({ key: element, text: element });
    });
    return _choice;
  }

  private _validdate(tabindex: number): boolean {
    let _errorCheck = this.state._errorCheck;
    if (tabindex == 0) {
        if (_errorCheck.firstTab.EmployeeDivision ) {
            _errorCheck.firstTab.showerror = true;
            this.setState({ _errorCheck: _errorCheck });
            return false;
        }
        else
            return true;
    }
    if(tabindex == 1){
      if ( _errorCheck.secondTab.ContractualStatus || _errorCheck.secondTab.Item_x0020_Description || _errorCheck.secondTab.Item_x0020_Value || _errorCheck.secondTab.ItemType || _errorCheck.secondTab.ItemGivenReceived || _errorCheck.secondTab.ItemIndividualCompany || _errorCheck.secondTab.ListIndividualCompany) {
        _errorCheck.secondTab.showerror = true;
        _errorCheck.secondTab.tabValid = false;
        this.setState({ _errorCheck: _errorCheck });
        return false;
    }
    else
    {
      _errorCheck.secondTab.tabValid = true;
      this.setState({ _errorCheck: _errorCheck });
        return true;
    }
    }
    return true;
  }

  private _OnInit = () => {
    
    let batch = sp.createBatch();
    sp.web.lists.getByTitle('Gifts Register').fields.getByInternalNameOrTitle('ItemType').inBatch(batch).get().then((res: Field) => {
      this.setState({ _itemTypeOption: this._formatChoices(res) });
    }).catch((e) => console.log(e));
    sp.web.lists.getByTitle('Gifts Register').fields.getByInternalNameOrTitle('ItemIndividualCompany').inBatch(batch).get().then((res: Field) => {
      this.setState({ _itemIndividualCompanyOption: this._formatChoices(res) });
    }).catch((e) => console.log(e));
    sp.web.lists.getByTitle('Gifts Register').fields.getByInternalNameOrTitle('ItemGivenReceived').inBatch(batch).get().then((res: Field) => {
      this.setState({ _itemGivenReceivedOption: this._formatChoices(res) });
    }).catch((e) => console.log(e));
    sp.web.lists.getByTitle('Gifts Register').fields.getByInternalNameOrTitle('ContractualStatus').inBatch(batch).get().then((res: Field) => {
      this.setState({ _contractualStatusOptions: this._formatChoices(res) });
    }).catch((e) => console.log(e));
    sp.web.lists.getByTitle('Gift Register Approval').items.select('Title','Approver/Id','ApproverLevel2/Id').top(20).orderBy('Title',true).expand('Approver','ApproverLevel2').get().then((res) => {
      let _choice: IDropdownOption[] = [];
      let _approverMatrix = new Map<string,any>();
      res.forEach(element => {
      _choice.push({ key: element["Title"], text: element["Title"] });
      _approverMatrix.set(element["Title"] , element)
    });
    this.setState({ _divisionOptions: _choice , _approverMatrix : _approverMatrix });
    }).catch((e) => console.log(e));
    sp.web.lists.getByTitle('Non Standard Approver').items.select('Requestor/Id','Approver1/Id','Approver2/Id').expand('Requestor','Approver2','Approver1').getAll().then((res) => {
      let _nonStandardApproverMatrix = new Map<string,any>();
      res.forEach(element => {
        _nonStandardApproverMatrix.set(element["Requestor"]["Id"] , element)
    });
    this.setState({  _nonStandardApproverMatrix : _nonStandardApproverMatrix });
    }).catch((e) => console.log(e));
    
    sp.web.currentUser.inBatch(batch).get().then(_ => this.setState({RequestedById : _.Id , _currentUserId : _.Id}))
    // sp.web.lists.getByTitle('Gifts Register').fields.filter('Hidden eq false and ReadOnlyField eq false').inBatch(batch).get().then(i => i.map((item) => console.log(item['StaticName'] +  '-' + item['TypeAsString']))).catch(e => console.log(e));
    batch.execute().then(itemDone =>{
    if(this.props.itemID) 
    {
      let _selectField: string[] = ['Requestor_x0020_Name',
      'Requestor_x0020_Email','Status','ID',
      'Employee_x0020_Title',
      'ManagerName',
      'RequestForSomeone','NonStandardApproval',
      'GiftDate',
      'ContractualStatus','RequestedBy/Title','Manager/Id','Approver1/Id','Approver2/Id','RequestedBy/Id','Requestor/Id','ActionedBy/Id',
      'ItemGivenReceived','ListIndividualCompany','Item_x0020_Government','IsAmountReasonable',
      'EmployeeDivision','Item_x0020_Value','ItemIndividualCompany','ItemType','Item_x0020_Description',
      'Attachments'];
      sp.web.lists.getByTitle('Gifts Register').items.getById(Number(this.props.itemID)).select(..._selectField).expand('ActionedBy','RequestedBy','Manager','Approver1','Approver2','Requestor').get().then(item => {
        console.log(item);
        this.setState({
          ID : item['ID'] != null ? item['ID'] : -1,
          Requestor_x0020_Name : item['Requestor_x0020_Name'] != null ? item['Requestor_x0020_Name'] : '',
          Requestor_x0020_Email : item['Requestor_x0020_Email'] != null ? item['Requestor_x0020_Email'] : '',
          Employee_x0020_Title : item['Employee_x0020_Title'] != null ? item['Employee_x0020_Title'] : '',
          ManagerName : item['ManagerName'] != null ? item['ManagerName'] : '',
          RequestForSomeone : item['RequestForSomeone'] != null ? item['RequestForSomeone'] : false,
          GiftDate : new Date(item['GiftDate']),
          ContractualStatus : item['ContractualStatus'] != null ? item['ContractualStatus'] : [],
          ItemGivenReceived : item['ItemGivenReceived'] != null ? item['ItemGivenReceived'] : '',
          EmployeeDivision : item['EmployeeDivision'] != null ? item['EmployeeDivision'] : '',
          Item_x0020_Value : item['Item_x0020_Value'] != null ? item['Item_x0020_Value'] : '',
          ItemIndividualCompany : item['ItemIndividualCompany'] != null ? item['ItemIndividualCompany'] : '',
          ItemType : item['ItemType'] != null ? item['ItemType'] : '',
          Item_x0020_Description : item['Item_x0020_Description'] != null ? item['Item_x0020_Description'] : '',
          ListIndividualCompany : item['ListIndividualCompany'] != null ? item['ListIndividualCompany'] : '',
          Item_x0020_Government: item['Item_x0020_Government'] != null ? item['Item_x0020_Government'] : false,
          IsAmountReasonable : item['IsAmountReasonable'] != null ? item['IsAmountReasonable'] : false,
          RequestedByTitle : item["RequestedBy"] !=null ? item["RequestedBy"]["Title"] : '',
          RequestedById : item["RequestedBy"] !=null ? item["RequestedBy"]["Id"] :null,
          ManagerId : item["Manager"] !=null ? item["Manager"]["Id"] : null,
          Approver1Id : item["Approver1"] !=null ? item["Approver1"]["Id"] : null,
          RequestorId : item["Requestor"] !=null ? item["Requestor"]["Id"] : null,
          Status : item["Status"] !=null ? item["Status"] : '',
          NonStandardApproval : item["NonStandardApproval"] !=null ? item["NonStandardApproval"] :false,
          
        })
        if (item["Attachments"]) {
          sp.web.lists.getByTitle("Gifts Register").items.getById(Number(this.props.itemID)).attachmentFiles.get().then((files: AttachmentFile[]) => {
              console.log("Attachment Fetched");
              let _AttachmentFiles = [];
              files.map((element, index) => {
                  let fileprop = {
                      FileName: element["FileName"],
                      URL: element["ServerRelativeUrl"]
                  };
                  _AttachmentFiles.push(fileprop);
              });
              this.setState({ AttachmentFiles: _AttachmentFiles });
          });
         
          
      }
      if(item["Approver2"]){
        let _approver2Id = [];
        item["Approver2"].map(_ => {
          _approver2Id.push(_.Id);
        });
        this.setState({Approver2Id : _approver2Id})
      }
      if(item["ActionedBy"]){
        let _actionedById = [];
        item["ActionedBy"].map(_ => {
          _actionedById.push(_.Id);
        });
        this.setState({ActionedById : _actionedById})
      }
      this._InitApproval();
      }).catch(e => console.log(e))
      
    }
    else{
      this._getUserInformation(this.props.context.pageContext.user.email);
    }
    }).catch(e => console.log(e));  
  }

  private _InitApproval = () => {

    //Show Approval to Manager
    if(this.state.Status == Status.Submitted && this.state._currentUserId == this.state.ManagerId)
      this.setState({_showApprover : true});
    //Show Approval to Approver 1
    else if(this.state.Status == Status.Approver1Pending && this.state._currentUserId == this.state.Approver1Id)
      this.setState({_showApprover : true});
    //Show Approval to Approver 2
    else if(this.state.Status == Status.Approver2Pending &&   this.state.Approver2Id.indexOf(this.state._currentUserId) > -1)
      this.setState({_showApprover : true});
    else
      console.log('User is not approver or item is already approved/rejected');
  }

  private _getUserInformation = (email : string) => {

    let _Requestor_x0020_Name = "";
      // let _Requestor_x0020_Phone = "";
      let _Requestor_x0020_Email = "";
      let _RequestorId = null
      let _Position = "";
      let _Title = "";
      let _ManagerNameLoginName = "";
      sp.web.siteUsers.getByEmail(email).get().then(e => {
        _RequestorId = e.Id
        sp.profiles.getPropertiesFor(e.LoginName).then((d) => {
        // console.log(d);
        d.UserProfileProperties.map((e) => {
          if(e.Key == "PreferredName")
            _Requestor_x0020_Name = e.Value;
          // if(e.Key == "WorkPhone")
          //   _Requestor_x0020_Phone = e.Value;
          if(e.Key == "WorkEmail")
            _Requestor_x0020_Email = e.Value;
          if(e.Key == "Title")
            _Title = e.Value;
          if(e.Key == "Department")
            _Position = e.Value;
          if(e.Key == "Manager")
            _ManagerNameLoginName = e.Value;
        });
        sp.web.siteUsers.getByLoginName(_ManagerNameLoginName).get().then((p)=>{
            this.setState({
              Requestor_x0020_Name : _Requestor_x0020_Name,
              Requestor_x0020_Email : _Requestor_x0020_Email,
              Employee_x0020_Title : _Title + ' | ' + _Position,
              ManagerName : p.Title,
              RequestorId : _RequestorId,
              ManagerId : p.Id
            });
        }).catch((e) => {console.log(e);this.setState({_managerError : true})});
      }).catch((e) => console.log(e));
      });
    
  }

  private _save = () => {
    if(this.state._errorCheck.secondTab.tabValid)
    {
    this.setState({_showSpinner : true});
    let fieldValues = {
      Requestor_x0020_Name : this.state.Requestor_x0020_Name,
      // Requestor_x0020_Phone : this.state.Requestor_x0020_Phone,
      Requestor_x0020_Email : this.state.Requestor_x0020_Email,
      Employee_x0020_Title : this.state.Employee_x0020_Title,
      EmployeeDivision : this.state.EmployeeDivision,
      GiftDate : this.state.GiftDate,
      Item_x0020_Value : this.state.Item_x0020_Value,
      ContractualStatus : { results: this.state.ContractualStatus },
      ItemGivenReceived : this.state.ItemGivenReceived,
      RequestedById : this.state.RequestedById,
      RequestForSomeone : this.state.RequestForSomeone,
      ManagerName : this.state.ManagerName,
      ItemIndividualCompany : this.state.ItemIndividualCompany,
      ItemType : this.state.ItemType,
      Item_x0020_Description : this.state.Item_x0020_Description,
      ListIndividualCompany : this.state.ListIndividualCompany,
      Item_x0020_Government : this.state.Item_x0020_Government,
      IsAmountReasonable : this.state.IsAmountReasonable,
      RequestorId : this.state.RequestorId,
      ManagerId :  this.state.ManagerId,
      NonStandardApproval : this.state._nonStandardApproverMatrix.has(this.state.RequestorId) ,
      Approver1Id :Number(this.state.Item_x0020_Value) > 250 ? this.state._nonStandardApproverMatrix.has(this.state.RequestorId) ? this.state._nonStandardApproverMatrix.get(this.state.RequestorId).Approver1.Id :  this.state.Approver1Id : null,
      Approver2Id :{results :  Number(this.state.Item_x0020_Value) > 250 ?  this.state._nonStandardApproverMatrix.has(this.state.RequestorId) ? this.state._nonStandardApproverMatrix.get(this.state.RequestorId).Approver2 ? [this.state._nonStandardApproverMatrix.get(this.state.RequestorId).Approver2.Id] : [] : this.state.Approver2Id : []},
      Status : Number(this.state.Item_x0020_Value) > 250 ? this.state._nonStandardApproverMatrix.has(this.state.RequestorId) ? Status.Approver1Pending : Status.Submitted : Status.Approved,
      CurrentApproverId : {results : Number(this.state.Item_x0020_Value) > 250  ? this.state._nonStandardApproverMatrix.has(this.state.RequestorId) ? [this.state._nonStandardApproverMatrix.get(this.state.RequestorId).Approver1.Id] : [this.state.ManagerId] : [] } 
    }
    console.log(fieldValues)
    sp.web.lists.getByTitle('Gifts Register').items.add({
      ...fieldValues
    }).then((result : ItemAddResult) => {
      console.log(result.item);
      this._addAttachments(result.data.ID) 

    });
  }
  else 
  this.setState({showValidationError : true})
  }

  private _addAttachments = (id: number) => {

    var element = document.getElementById("upload");
    var files = element["files"];
    if (files.length > 0) {
      let _files: AttachmentFileInfo[] = [];
      for (let i = 0; i < files.length; i++) {
        _files.push({ name: files[i].name, content: files[i] });
      }
      sp.web.lists.getByTitle('Gifts Register').items.getById(id).attachmentFiles.addMultiple(_files)
      .then(_ => {
        sp.web.lists.getByTitle('Gifts Register').items.getById(id).update({
        View : {
          "__metadata": { type: "SP.FieldUrlValue" },
          Description: `GBE Request - ${id}`,
          Url: `${this.props.context.pageContext.web.absoluteUrl}/SitePages/GBE-Request.aspx?item=${id} `
      }
      }).then( _ => document.location.href = this.props.context.pageContext.web.absoluteUrl).catch(e => console.log(e))})
        .catch(e => console.log(e));
    }
  
    else {
        console.log("No Files to upload");
        sp.web.lists.getByTitle('Gifts Register').items.getById(id).update({
          View : {
            "__metadata": { type: "SP.FieldUrlValue" },
            Description: `GBE Request - ${id}`,
            Url: `${this.props.context.pageContext.web.absoluteUrl}/SitePages/GBE-Request.aspx?item=${id} `
        }
        }).then( _ => document.location.href = this.props.context.pageContext.web.absoluteUrl).catch(e => console.log(e));
    }

}

  private _contractualStatusChanged = (_,v) => {
    let _contractualStatus = this.state.ContractualStatus;
    if (v) {
      _contractualStatus.push(_.target.id)
    }
    else {
      this.state.ContractualStatus.map((item, index) => {
        if (item == _.target.id) {
          _contractualStatus.splice(index, 1);
        }
      });
    }
    let _errorCheck = this.state._errorCheck;
    _contractualStatus.length > 0 ? _errorCheck.secondTab.ContractualStatus = false : _errorCheck.secondTab.ContractualStatus = true;
    this.setState({ ContractualStatus: _contractualStatus , _errorCheck : _errorCheck });
  }

  private _contractStatusChoices = (_disabled : boolean) : JSX.Element[]=> {
    let element : JSX.Element[] = [] ;
    this.state._contractualStatusOptions.forEach((value,index) => {
      let item : JSX.Element = <Checkbox id={value.key} label={value.key} disabled={_disabled} checked={this.state.ContractualStatus.indexOf(value.key) > -1 ? true : false } onChange={(e,v) => this._contractualStatusChanged(e,v)}/>;
      element.push(item) ;
      element.push(<br/>);
    })
    return element;
  }

  private _getRequesterValue = (_: any[]) => {
    if(_.length > 0)  
      this._getUserInformation(_[0].secondaryText);
  }

  public _tabClick(item: PivotItem): void {

    // const tabIndex: number = Number(item.props.itemKey);
    if(this.props.itemID){
      this.setState({ _selectedKey: Number(item.props.itemKey) })
    }
    else if (this._validdate(this.state._selectedKey)){   
      this.setState({ _selectedKey: Number(item.props.itemKey) })
    }
    else{
        this.setState({ _selectedKey : this.state._selectedKey})
    }
  
  }

  private _onValidFieldTabValidation = (id ,value:string) => {
    let tab = ''
    switch(this.state._selectedKey){
        case 0 : tab = 'firstTab';break;
        case 1 : tab ='secondTab' ; break;
        case 3 : tab = 'approverTab';break;
    }
    let _errorCheck = this.state._errorCheck;
    value.length > 0 ? _errorCheck[tab][id] = false : _errorCheck[tab][id] = true;
    if(id == 'EmployeeDivision')
      this.setState({ [id]: value, _errorCheck: _errorCheck , Approver1Id : this.state._approverMatrix.get(value).Approver ? this.state._approverMatrix.get(value).Approver.Id : null , Approver2Id : this.state._approverMatrix.get(value).ApproverLevel2.map( _ => _.Id ) })
    else  
      this.setState({ [id]: value, _errorCheck: _errorCheck })
}

private _onFormatDate = (date: Date): string => {
  return date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear());
}

private _validateAmount = (_) => {
  console.log(Number(_))
  if (isNaN(_))
    _ = _.slice(0, -1);
  else if(Number(_) < 1)
     _ = _.slice(0, -1);

  let _errorCheck = this.state._errorCheck;
  _errorCheck.secondTab.Item_x0020_Value = _.length > 0 ? false : true
  this.setState({ Item_x0020_Value: _, _errorCheck: _errorCheck });
}

  public render(): React.ReactElement<IGiftRegisterProps> {
    console.log(this.state);
    let _disabled = this.props.itemID ? true : false;
    let _errorMsg = 'Field can not be blank'
    return (
      
      <div>
        {/* <h1>Brookfield Hospitality / Gift Approval Form</h1> */}
        <Pivot linkSize={PivotLinkSize.normal} styles={{ root: { display: 'flex', flexWrap: 'wrap' } }} linkFormat={PivotLinkFormat.links} onLinkClick={this._tabClick} selectedKey={`${this.state._selectedKey}`}>
        <PivotItem headerText="Requestor Information" itemKey="0">
          <Toggle label="Request on behalf of someone ?" disabled={_disabled} onText="Yes" offText="No" checked={this.state.RequestForSomeone}
          onChange={(e,v) => {this.setState({RequestForSomeone : v});!v ? this._getUserInformation(this.props.context.pageContext.user.email) : ''}}></Toggle>

          {this.state.RequestForSomeone && !_disabled && <PeoplePicker context={this.props.context} disabled={_disabled} titleText="Request For" personSelectionLimit={1} groupName={""} showtooltip={true} showHiddenInUI={false} 
            principalTypes={[PrincipalType.User]} resolveDelay={200} ensureUser={true} defaultSelectedUsers={[]} selectedItems={this._getRequesterValue} />}

          {this.state.RequestForSomeone && _disabled && <TextField label="Requested By" disabled={_disabled} value={this.state.RequestedByTitle} />}
          <TextField label="Requestor" disabled value={this.state.Requestor_x0020_Name}></TextField>
          {/* <TextField label="Requestor Phone" disabled value={this.state.Requestor_x0020_Phone}></TextField> */}
          <TextField label="Requestor Email" disabled value={this.state.Requestor_x0020_Email}></TextField>
          <TextField label="Position Title" disabled value={this.state.Employee_x0020_Title}></TextField>
          <Dropdown label="Division" disabled={_disabled} selectedKey={this.state.EmployeeDivision} errorMessage={this.state._errorCheck.firstTab.EmployeeDivision && this.state._errorCheck.firstTab.showerror ? _errorMsg : undefined} required options={this.state._divisionOptions} onChange={(e,v) => this._onValidFieldTabValidation('EmployeeDivision' , v.text)} ></Dropdown>
          <TextField label="Manager Name" disabled value={this.state.ManagerName}></TextField>
        </PivotItem>
        <PivotItem headerText="Additional Information" itemKey="1" onRenderItemLink={this._customRenderer} >
          <div style={{float:'left' , width:'45%'}}>
          <DatePicker strings={DayPickerStrings} formatDate={this._onFormatDate} disabled={_disabled} label="Date of Gift or Entertainment Received/or Given" value={this.state.GiftDate} onSelectDate={(d) => this.setState({GiftDate : d})}></DatePicker>
          <TextField label="Estimated value per individual (AUD)"  required errorMessage={this.state._errorCheck.secondTab.Item_x0020_Value && this.state._errorCheck.secondTab.showerror ? _errorMsg : undefined} disabled={_disabled}   value={this.state.Item_x0020_Value} onChange={(e,v) => this._validateAmount(v)}></TextField>
          <Dropdown label="Category" disabled={_disabled} required errorMessage={this.state._errorCheck.secondTab.ItemType && this.state._errorCheck.secondTab.showerror ? _errorMsg : undefined} defaultSelectedKey={this.state.ItemType}  options={this.state._itemTypeOption} onChange={(e,v) => this._onValidFieldTabValidation('ItemType' , v.text)} ></Dropdown>
          <Dropdown label="Was this…." required errorMessage={this.state._errorCheck.secondTab.ItemGivenReceived && this.state._errorCheck.secondTab.showerror ? _errorMsg : undefined} defaultSelectedKey={this.state.ItemGivenReceived} disabled={_disabled}  options={this.state._itemGivenReceivedOption} onChange={(e,v) => this._onValidFieldTabValidation('ItemGivenReceived' , v.text)}></Dropdown>
          {this.state.ItemGivenReceived == 'Given' && <Toggle disabled={_disabled} checked={this.state.Item_x0020_Government} defaultValue={this.state.ItemGivenReceived}  label="Was this Gift, Benefit and Entertainment given to Public Official?" onText="Yes" offText="No" onChange={(e,v) =>this.setState({Item_x0020_Government : v})}></Toggle>}
          <Toggle  disabled={_disabled} onChange={(e,v) =>this.setState({IsAmountReasonable : v})} checked={this.state.IsAmountReasonable} label="Is the amount reasonable compared to the volume of affairs with 3rd party?" onText="Yes" offText="No" />          
            <Label required >Status of the contractual relationship with 3rd party?</Label>
            {this.state._errorCheck.secondTab.ContractualStatus && this.state._errorCheck.secondTab.showerror && <p style={{paddingTop:'5px',color:'rgb(168, 0, 0)',fontWeight : 400, fontSize:'12px'}}><span>Please choose one or more options from the below selection</span></p>}<br/>
            {this._contractStatusChoices(_disabled)}
            </div>
          <div style={{width:"1px", height:"450px",margin:"5%",background:"#a6a6a6",float:"left"}}></div>
          <div style={{float:'left' , width:'44.7%'}}>
          <Dropdown disabled={_disabled} required errorMessage={this.state._errorCheck.secondTab.ItemIndividualCompany && this.state._errorCheck.secondTab.showerror ? _errorMsg : undefined}  label="Does this Gift, Benefit or Entertainment relate to..." defaultSelectedKey={this.state.ItemIndividualCompany} options={this.state._itemIndividualCompanyOption} onChange={(e,v) => this._onValidFieldTabValidation('ItemIndividualCompany' , v.text)}></Dropdown>            
            <TextField disabled={_disabled} required errorMessage={this.state._errorCheck.secondTab.ListIndividualCompany && this.state._errorCheck.secondTab.showerror ? _errorMsg : undefined}  label="List of individuals including Brookfield employees and third parties (incl. company name)" multiline rows={7} defaultValue={this.state.ListIndividualCompany} onChange={(e,v) => this._onValidFieldTabValidation('ListIndividualCompany' , v)} ></TextField>            
            <TextField required  disabled={_disabled} errorMessage={this.state._errorCheck.secondTab.Item_x0020_Description && this.state._errorCheck.secondTab.showerror ? _errorMsg : undefined}  label="Description (incl. the reason and type of Gift, Benefit or Entertainment given or received)" multiline rows={10} defaultValue={this.state.Item_x0020_Description} onChange={(e,v) => this._onValidFieldTabValidation('Item_x0020_Description' , v)}></TextField>
            </div>
        </PivotItem>
        <PivotItem headerText="Supporting Documents" itemKey="2" >
          <div>
          
          <TextField  disabled={_disabled}  type="file" id="upload" label="Additional Supporting Documentation?" borderless multiple></TextField><br/>
          <p style={{paddingTop:'5px',color:'rgb(168, 0, 0)',fontWeight : 400, fontSize:'12px'}}><span>Note : You can upload multiple files using Ctrl or Shift key</span></p>
          {this._showAttachemnt()}
          </div>
          <div style={{margin:"40px" , textAlign:"center"}}>
          {!this.state._showSpinner && <>
            {!_disabled &&<PrimaryButton style={{margin:"0px 20px"}} onClick={this._save}  >Submit</PrimaryButton>}
            <Button style={{margin:"0px 20px"}} onClick={_ => document.location.href = this.props.context.pageContext.web.absoluteUrl}>Cancel</Button>
            </>
          }
          {this.state._showSpinner && <>
          <Label>Please wait while your data is being saved....</Label><br/>
          <Spinner size={SpinnerSize.large} style={{margin:"0px 20px"}} /></>}
          </div>
        </PivotItem>
        <PivotItem headerText="Approve or Reject" itemKey="3"  onRenderItemLink={this._customRenderer} headerButtonProps={{show : this.state._showApprover}}>
            <TextField label="Comments" onChange={(e,v) => this._onValidFieldTabValidation('comment' , v)} multiline rows={6} errorMessage={this.state._errorCheck.approverTab.comment && this.state._errorCheck.approverTab.showerror ? _errorMsg : undefined} required></TextField>
            <div style={{margin:"40px" , textAlign:"center"}}>
              <PrimaryButton style={{margin:"0px 20px",background :"#0f3557"}} onClick={this._approve} >Approve</PrimaryButton>
              <PrimaryButton style={{margin:"0px 20px",background:"#4d4d4d"}} onClick={this._reject}>Reject</PrimaryButton>
            </div>
        </PivotItem>
        </Pivot>
        <Dialog hidden={!this.state._managerError} dialogContentProps={{ type: DialogType.normal, title: 'Error', subText: 'We are unable to populate your manager information. Please contact complianceau@brookfield.com for assistance.' }}
                    modalProps={{ isBlocking: true, styles: { main: { maxWidth: 450 } } }}>
                    <DialogFooter>
                        <PrimaryButton onClick={() => document.location.href = this.props.context.pageContext.web.absoluteUrl} text="Close" />
                    </DialogFooter>
                </Dialog>
      </div>
    );
  }



  private _approve = () => {
    if(this.state._errorCheck.approverTab.comment)
    {
      let _errorCheck = this.state._errorCheck;
      _errorCheck.approverTab.showerror = true;
      this.setState({_errorCheck : _errorCheck});
    }
    else{
      //Approved by Manager
      let _actionedById = this.state.ActionedById;
      if(this.state.Status == Status.Submitted && this.state._currentUserId == this.state.ManagerId)
      {
        _actionedById.push(this.state.ManagerId);
        sp.web.lists.getByTitle('Gifts Register').items.getById(this.state.ID).update({
          Status : this.state.Approver1Id ? Status.Approver1Pending : Status.Approver2Pending,
          CurrentApproverId : {results :this.state.Approver1Id ? [this.state.Approver1Id] : this.state.Approver2Id},
          ActionedById : {results : _actionedById},
          ManagerComment : this.state.comment 
        }).then(_ => document.location.href = this.props.context.pageContext.web.absoluteUrl).catch(_ => console.log(_))
      }
      //Approved By Approver 1
      else if(this.state.Status == Status.Approver1Pending && this.state._currentUserId == this.state.Approver1Id)
      {
        _actionedById.push(this.state.Approver1Id); 
        sp.web.lists.getByTitle('Gifts Register').items.getById(this.state.ID).update({
          Status : this.state.Approver2Id.length > 0 ? Status.Approver2Pending : Status.Approved,
          CurrentApproverId : {results :this.state.Approver2Id.length > 0 ? this.state.Approver2Id : [] },
          ActionedById : {results : _actionedById},
          Approver1Comment : this.state.comment 
        }).then(_ => document.location.href = this.props.context.pageContext.web.absoluteUrl).catch(_ => console.log(_))
      }
      //Approved by Approver 2
      else if(this.state.Status == Status.Approver2Pending && this.state.Approver2Id.indexOf(this.state._currentUserId) > -1)
      {
        _actionedById = _actionedById.concat(this.state.Approver2Id); 
        sp.web.lists.getByTitle('Gifts Register').items.getById(this.state.ID).update({
          Status : Status.Approved,
          Approver2Comment : this.state.comment,
          ActionedById : {results : _actionedById},
          CurrentApproverId : {results :[]},
        }).then(_ => document.location.href = this.props.context.pageContext.web.absoluteUrl).catch(_ => console.log(_))
      }
    }
  }

  private _reject = () => {
    if(this.state._errorCheck.approverTab.comment)
    {
      let _errorCheck = this.state._errorCheck;
      _errorCheck.approverTab.showerror = true;
      this.setState({_errorCheck : _errorCheck});
    }
    else{
      //Rejected by Manager
      let _actionedById = this.state.ActionedById;
      if(this.state.Status == Status.Submitted && this.state._currentUserId == this.state.ManagerId)
      {
        _actionedById.push(this.state.ManagerId);
        sp.web.lists.getByTitle('Gifts Register').items.getById(this.state.ID).update({
          Status : Status.RejectedManager,
          ManagerComment : this.state.comment ,
          ActionedById : {results :_actionedById},
          CurrentApproverId : {results :[]},
        }).then(_ => document.location.href = this.props.context.pageContext.web.absoluteUrl).catch(_ => console.log(_))
      }
      //Rejected By Approver 1
      else if(this.state.Status == Status.Approver1Pending && this.state._currentUserId == this.state.Approver1Id)
      {
        _actionedById.push(this.state.Approver1Id); 
        sp.web.lists.getByTitle('Gifts Register').items.getById(this.state.ID).update({
          Status : Status.RejectedApprover1,
          Approver1Comment : this.state.comment ,
          CurrentApproverId : {results :[]},
          ActionedById : {results : _actionedById},
        }).then(_ => document.location.href = this.props.context.pageContext.web.absoluteUrl).catch(_ => console.log(_))
      }
      //Rejected by Approver 2
      else if(this.state.Status == Status.Approver2Pending && this.state.Approver2Id.indexOf(this.state._currentUserId) > -1)
      {
        _actionedById = _actionedById.concat(this.state.Approver2Id); 
        sp.web.lists.getByTitle('Gifts Register').items.getById(this.state.ID).update({
          Status : Status.RejectedApprover2,
          Approver2Comment : this.state.comment ,
          CurrentApproverId : {results :[]},
          ActionedById : {results : _actionedById},
        }).then(_ => document.location.href = this.props.context.pageContext.web.absoluteUrl).catch(_ => console.log(_))
      }
    }
  }


  private _customRenderer(link: IPivotItemProps, defaultRenderer: (link: IPivotItemProps) => JSX.Element): JSX.Element {
    let element: JSX.Element ;
    switch(link.headerText)
    {
      case "Approve or Reject" : 
      element = link.headerButtonProps.show ? <span>{defaultRenderer(link)}</span> :  <span style={{display:"none"}}>{defaultRenderer(link)}</span>; 
      break;
      case "Additional Information" : 
      element = !this.state._errorCheck.secondTab.tabValid && this.state.showValidationError ? <span>{defaultRenderer(link)} <Icon iconName="Error" style={{ color: 'red' }}/></span> :  <span>{defaultRenderer(link)}</span>;
      break;
    }
    return element;
  }

  private _showAttachemnt = (): JSX.Element[] => {
    let linkelement: JSX.Element[] = [];
    this.state.AttachmentFiles.map((element, index) => {
        linkelement.push(<li><Link href={element.URL}>{element.FileName}</Link></li>);
    });
    return linkelement;
}
}
