import * as React from 'react';
import { ISharePointFormColumns } from './ISharePointListColumns';

export const handleSingleSelectedPeoplePicker=(items:any[],setFormData:React.Dispatch<React.SetStateAction<ISharePointFormColumns>>)=>{
if(items.length>0){
    setFormData(prev=>({...prev,Admin:items[0].text,AdminId:items[0].id}))
}
else{
    setFormData(prev=>({...prev,Admin:"",AdminId:0}));
}

}


export const handleMultiSelectedPeoplePicker=(items:any[],setFormData:React.Dispatch<React.SetStateAction<ISharePointFormColumns>>)=>{
setFormData(prev=>({...prev,Manager:items.map(i=>i.text),ManagerId:items.map(i=>i.id)}));

}