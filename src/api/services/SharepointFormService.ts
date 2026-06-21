import {Web} from "@pnp/sp/presets/all";
import { ListNames } from "../../enum/ListNames";
import { ISharePointFormColumns } from "../../models/ISharePointListColumns";

export class SharePointFormServiceClass{
private web;
constructor(siteurl:string){
    this.web=Web(siteurl);
}

public async createItems(FormData:ISharePointFormColumns):Promise<any>{

    try{
const list=this.web.lists.getByTitle(ListNames.FirstList); //it will hold the listname
// siteurl/web/lists/getbytitle()/items.add()
const result=await list.items.add({
    Title:FormData.Name,
    EmailAddress:FormData.Email,
    Age:parseInt(FormData.Age),
    Salary:parseFloat(FormData.Compensation),
    Score:FormData.Experience,
    Permission:FormData.Permission,
    Address:FormData.Address,
    AdminId:FormData.AdminId,
    ManagerId:{results:FormData.ManagerId}
})
return result;
    }
    catch(err){
console.log(`Error while creating the items`,err);

    }
    finally{
console.log(`I am free will`);
    }
}

}