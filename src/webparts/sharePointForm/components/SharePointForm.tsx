import * as React from 'react';
// import styles from './SharePointForm.module.scss';
import type { ISharePointFormProps } from './ISharePointFormProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { SharePointFormServiceClass } from '../../../api/services/SharepointFormService';
import { ISharePointFormColumns } from '../../../models/ISharePointListColumns';
import { Dialog } from '@microsoft/sp-dialog';
import { ChoiceGroup, Dropdown, PrimaryButton, Slider, TextField, Toggle } from '@fluentui/react';
import {  PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { handleMultiSelectedPeoplePicker, handleSingleSelectedPeoplePicker } from '../../../models/PeoplePickerHandler';
import { handleSkillsChange } from '../../../models/MultiselectDropdownModel';
const SharePointForm:React.FC<ISharePointFormProps>=(props)=>{
  const [formdata,setFormdata]=React.useState<ISharePointFormColumns>({
    Name:"",
    Email:"",
    Age:"",
    Compensation:"",
    Experience:0,
    Address:"",
    Permission:false,
    Manager:[],
    ManagerId:[],
    Admin:"",
    AdminId:0,
    Skills:[],
    Department:"",
    City:"",
    Gender:""
  });

  const addItems=async()=>{
    try{
const _service=new SharePointFormServiceClass(props.siteurl);
const result =await _service.createItems(formdata);
Dialog.alert(`Item added successfully with id ${result.data.id}`);

setFormdata({
   Name:"",
    Email:"",
    Age:"",
    Compensation:"",
    Experience:0,
    Address:"",
    Permission:false,
     Manager:[],
    ManagerId:[],
    Admin:"",
    AdminId:0,
    Skills:[],
    Department:"",
    City:"",
    Gender:""
})
//I am master
    }
    catch(err){
console.log(err);
    }
  }

  //event handling 
   const handleChange=React.useCallback((field:keyof ISharePointFormColumns,value:string|number|boolean)=>{
          setFormdata(prev=>({...prev,[field]:value}))
      },[]);
  return(
    <>
           <TextField
           label="Name"
           value={formdata.Name}
           onChange={(e,newValue)=>handleChange("Name",newValue||"")}
           />
            <TextField
           label="Email Address"
           value={formdata.Email}
           onChange={(e,newValue)=>handleChange("Email",newValue||"")}
           />
            <TextField
           label="Age"
           value={formdata.Age}
           onChange={(e,newValue)=>handleChange("Age",newValue||"")}
           />
           
            <TextField
           label="Compensation"
           value={formdata.Compensation}
           onChange={(e,newValue)=>handleChange("Compensation",newValue||"")}
           prefix="₹"
           suffix="per annum"
           />
           <Slider
           label="Experience"
           min={0}
           max={20}
           value={formdata.Experience}
           onChange={(value)=>handleChange("Experience",value)}
           />
           <Toggle
           label="Permission"
           checked={formdata.Permission}
            onChange={(e,newValue)=>handleChange("Permission",!!newValue)}
           />
           <PeoplePicker
    context={props.context as any}
    titleText="Admin"
    personSelectionLimit={1}
    showtooltip={true}
ensureUser={true}
    onChange={(items)=>handleSingleSelectedPeoplePicker(items,setFormdata)}
  defaultSelectedUsers={[formdata.Admin?formdata.Admin:""]}
    principalTypes={[PrincipalType.User]}
    resolveDelay={1000}
    webAbsoluteUrl={props.siteurl} />

         <PeoplePicker
    context={props.context as any}
    titleText="Manager"
    personSelectionLimit={3}
    showtooltip={true}
ensureUser={true}
    onChange={(items)=>handleMultiSelectedPeoplePicker(items,setFormdata)}
  defaultSelectedUsers={formdata.Manager}
    principalTypes={[PrincipalType.User]}
    resolveDelay={1000}
    webAbsoluteUrl={props.siteurl} />
    {/* Department */}
    <Dropdown
    label='Department'
    placeholder='--select--'
    options={props.singleselecteddropdown}
    selectedKey={formdata.Department}
    onChange={(_,options)=>handleChange("Department",options?.key as string)}
    />
    {/* City */}
    <Dropdown
    label='City'
    placeholder='--select--'
    options={props.lookupval}
    selectedKey={formdata.City}
    onChange={(_,options)=>handleChange("City",options?.key as string)}
    />
    {/* Raido Button -Gender */}

    <ChoiceGroup
    label='Gender'
    options={props.radiooption}
     selectedKey={formdata.Gender}
    onChange={(_,options)=>handleChange("Gender",options?.key as string)}
    />
    {/* multiselecy */}
     <Dropdown
    label='Skills'
    placeholder='--select--'
    options={props.multiselectdropdown}
    defaultSelectedKeys={formdata.Skills}
    multiSelect
    onChange={(_,opt)=>handleSkillsChange(opt!,formdata,setFormdata)}
    />
          <TextField
           label="Address"
           value={formdata.Address}
           onChange={(e,newValue)=>handleChange("Address",newValue||"")}
           multiline
           rows={5}
           />
           <br/>
           <PrimaryButton
           text='Save'
           onClick={addItems}
           iconProps={{iconName:'save'}}
           />
           </>
  )
}
export  default SharePointForm;
