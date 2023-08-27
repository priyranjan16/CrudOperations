import * as React from 'react';
import styles from './CrudOperations.module.scss';
import { ICrudOperationsProps } from './ICrudOperationsProps';
import { ICrudOperationsStates } from './ICrudOperationState';
import {sp,Web} from '@pnp/sp/presets/all';
import "@pnp/sp/items";
import "@pnp/sp/lists";
import {PeoplePicker,PrincipalType} from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { DatePicker, IDatePickerStrings, Label, PrimaryButton, TextField } from '@fluentui/react';
export default class CrudOperations extends React.Component<ICrudOperationsProps, ICrudOperationsStates> {
 constructor(props:any){
  super(props);
  sp.setup({
    spfxContext:this.props.context as any
  });
  this.state={
    IListItems:[],
    Title:"",
    Age:"",
    Address:"",
    Manager:"",
    ManagerId:"",
    JoiningDate:"",
    PhoneNumber:"",
    ID:0,
    HTML:[],
    ValidationErrors:{}
  }
 }
 public async componentDidMount() {
   await this.FetchData();
 }
 //Fetch Data
 public async FetchData(){
  const web=Web(this.props.siteurl);
  const items:any[]=await web.lists.getByTitle("CrudOperations").items.select("*","Manager/Title").expand("Manager").getAll();
  this.setState({IListItems:items});
  let html=await this.getHTML(items);
  this.setState({HTML:html});
 }

 //FindData

 public findData=(id:any):void=>{
  var itemID=id;
  var allItems=this.state.IListItems;
  var allItemsLength=allItems.length;
  if(allItemsLength>0){
    for(var i=0;i<allItems.length;i++){
      if(itemID===allItems[i].Id){
        this.setState({
          ID:itemID,
          Title:allItems[i].Title,
          Age:allItems[i].Age,
          Address:allItems[i].Address,
          PhoneNumber:allItems[i].PhoneNumber,
          Manager:allItems[i].Manager.Title,
          ManagerId:allItems[i].ManagerId,
          JoiningDate:new Date(allItems[i].JoiningDate)
        });
      }
    }
  }
 }
 //Get HTML Table
public async getHTML(items:any){
var tabledata=<table className={styles.table}>
<thead>
  <tr>
    <th>Employee Name</th>
    <th>Age</th>
    <th>Employee Address</th>
    <th>Phone Number</th>
    <th>Manager</th>
    <th>Joining Date</th>
  </tr>
</thead>
<tbody>
  {items && items.map((item:any,i:any)=>{
return[
  <tr key={i} onClick={()=>this.findData(item.ID)}>
    <td>{item.Title}</td>
    <td>{item.Age}</td>
    <td>{item.Address}</td>
    <td>{item.PhoneNumber}</td>
    <td>{item.Manager.Title}</td>
    <td>{FormatDate(item.JoiningDate)}</td>

  </tr>
]
  })}
</tbody>
</table>
return await tabledata;
}

//Validate Form Fields
private validateFormFields():boolean{
  const{Title,Address,PhoneNumber,Age}=this.state;
  const errors:any={};
  if(!Title){
    errors.EmployeeName="employee name can not be empty";
  }
  if(!Address||Address.length<15){
    errors.Address="Please write atleast 15 characters";
  }
  if(!PhoneNumber||!/^\d{10}$/.test(PhoneNumber.toString())){
    errors.PhoneNumber="Please enter 10 digits valid phonenumebr";
  }
  if(Age<21){
    errors.Age="you are not eligible age should greater than or equal to 21";
  }
  this.setState({ValidationErrors:errors});
  return Object.keys(errors).length===0;
}

 //Create Data
 public async saveData(){
  if(!this.validateFormFields()){
    return
  }
  //Co
  const web=Web(this.props.siteurl);
  await web.lists.getByTitle("CrudOperations").items.add({
    Title:this.state.Title,
    Age:this.state.Age,
    Address:this.state.Address,
    JoiningDate:this.state.JoiningDate,
    PhoneNumber:this.state.PhoneNumber,
    ManagerId:this.state.ManagerId
  })
  .then((data)=>{
    console.log("No error found");
    return data;
  })
  .catch((error)=>{
    console.error("Error found");
    throw error;
  });
  alert("data has been successfully addedd");
  this.setState({
    Title:"",
    Age:"",
    Address:"",
    JoiningDate:"",
    PhoneNumber:"",
    Manager:""
  });
  await this.FetchData(); //refresh 
 }

 //update Items
 public async updateData(){
  const web=Web(this.props.siteurl);
  await web.lists.getByTitle("CrudOperations").items.getById(this.state.ID).update({
    Title:this.state.Title,
    Age:this.state.Age,
    Address:this.state.Address,
    JoiningDate:this.state.JoiningDate,
    PhoneNumber:this.state.PhoneNumber,
    ManagerId:this.state.ManagerId
  })
  .then((data)=>{
    console.log("No error found");
    return data;
  })
  .catch((error)=>{
    console.error("Error found");
    throw error;
  });
  alert("data has been successfully updated");
  this.setState({
    Title:"",
    Age:"",
    Address:"",
    JoiningDate:"",
    PhoneNumber:"",
    Manager:""
  });
   await this.FetchData(); //refresh 
 }
 public async deleteData(){
  const web=Web(this.props.siteurl);
  await web.lists.getByTitle("CrudOperations").items.getById(this.state.ID).delete()
  .then((data)=>{
    console.log("No error found");
    return data;
  })
  .catch((error)=>{
    console.error("Error found");
    throw error;
  });
  alert("data has been successfully deleted");
  this.setState({
Title:"",
    Age:"",
    Address:"",
    JoiningDate:"",
    PhoneNumber:"",
    Manager:""
  });
  await this.FetchData(); //refresh 
 }
 //Form Event
 private handleChange=(fieldName:keyof ICrudOperationsStates,value:string|number|boolean):void=>{
if(fieldName==="PhoneNumber" ){
  value=value.toString().replace(/\D/g,"");
}
else if(fieldName==="Age"){
  value=value.toString().replace(/\D/g,"");
}
this.setState({[fieldName]:value}as unknown as Pick<ICrudOperationsStates,keyof ICrudOperationsStates>);
 }
 //Get PeoplePiker
private getPeoplepicker=(items:any[]):void=>{
if(items.length>0){
  this.setState({
    Manager:items[0].text,
    ManagerId:items[0].id
  });
}
else{
  this.setState({
    Manager:"",
    ManagerId:""
  })
}
}
  public render(): React.ReactElement<ICrudOperationsProps> {
 const{ValidationErrors}=this.state;

    return (
   <>
   {this.state.HTML}
   <div className={styles.btngroup}>
    <div>
      <PrimaryButton text="Save" iconProps={{iconName:"save"}} onClick={()=>this.saveData()}/>
    </div>
    <div>
      <PrimaryButton text="Update" iconProps={{iconName:"edit"}} onClick={()=>this.updateData()}/>
    </div>
    <div>
      <PrimaryButton text="Delete" iconProps={{iconName:"delete"}} onClick={()=>this.deleteData()}/>
    </div>
   </div>
   <div>
    <form>
      <div>
        <Label required>Employee Name:</Label>
        <TextField type="text" value={this.state.Title}
        onChange={(_,value)=>this.handleChange("Title",value||"")} 
errorMessage={ValidationErrors.EmployeeName}
iconProps={{iconName:"contact"}}
        
        />
      </div>
      <div>
      <Label required>Age:</Label>
        <TextField  value={this.state.Age?.toString()}
        onChange={(_,value)=>this.handleChange("Age",parseInt(value||""))} 
errorMessage={ValidationErrors.Age}
// iconProps={{iconName:"contact"}}
        
        />
      </div>
      <div>
      <Label required>Employee Address:</Label>
        <TextField type="text" value={this.state.Address}
        onChange={(_,value)=>this.handleChange("Address",value||"")} 
errorMessage={ValidationErrors.EmployeeAdress}
iconProps={{iconName:"location"}}
multiline
rows={5}
        
        />
      </div>
      <div>
      <Label required>Phone Number:</Label>
        <TextField  value={this.state.PhoneNumber?.toString()}
        onChange={(_,value)=>this.handleChange("PhoneNumber",parseInt(value||""))} 
errorMessage={ValidationErrors.PhoneNumber}
iconProps={{iconName:"phone"}}
        
        />
      </div>
      <div>
        <Label>Manager:</Label>
        <PeoplePicker context={this.props.context as any}
        personSelectionLimit={1}
        showtooltip={true}
        ensureUser={true}
        resolveDelay={1000}
        principalTypes={[PrincipalType.User]}
        defaultSelectedUsers={[this.state.Manager?this.state.Manager:""]}
        onChange={this.getPeoplepicker}
        />
      </div>
      <div>
        <Label>Joining Date:</Label>
        <DatePicker
        
        maxDate={new Date()}
        allowTextInput={false}
        strings={DatePickerStrings}
        value={this.state.JoiningDate}
        onSelectDate={(e)=>{this.setState({JoiningDate:e})}}
        aria-label='select a date' formatDate={FormatDate}
        />
      </div>
    </form>
   </div>
   </>
    );
  }
}
export const DatePickerStrings:IDatePickerStrings={
  months:["January","February","March","April","May","June","July","August","September","October","Novemebr","December"],
  shortMonths:["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],
  days:["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"],
  shortDays:["Sun","Mon","Tue","Wed","Thu","Fri","Sat"],
  goToToday:"Go To Today",
  prevMonthAriaLabel:"Go To Previous month",
  nextMonthAriaLabel:"Go To next Month",
  prevYearAriaLabel:"Go to Previous year",
  nextYearAriaLabel:"Go To next year",
  invalidInputErrorMessage:"invalid date format"
  
}

export const FormatDate=(date:any):string=>{
var date1=new Date(date);
var year=date1.getFullYear();
var month=(1+date1.getMonth()).toString();
month=month.length>1?month:'0'+month;
var day=date1.getDate().toString();
day=day.length>1?day:'0'+day;
return month +'/'+day+'/'+year;

}