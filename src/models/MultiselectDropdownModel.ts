import { ISharePointFormColumns } from "./ISharePointListColumns";
import { IDropdownOption } from "@fluentui/react";

export const handleSkillsChange=(options:IDropdownOption,formData:ISharePointFormColumns,setFormData:React.Dispatch<React.SetStateAction<ISharePointFormColumns>>)=>{
    const selectedKey=options.selected?[...formData.Skills,options?.key as string]:formData.Skills?.filter((key:any)=>key!==options.key);
    setFormData(prev=>({...prev,Skills:selectedKey}))
}