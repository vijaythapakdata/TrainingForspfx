export interface ISharePointListColumns{
    Name:string;
    EmailAddress:string;
    PhoneNumber:string;
    Compensation:number;
    Age:number;
    FullAddress?:string|any|undefined;
    Salary:any;
    City:{
        Title:string;
        Id:number
    };
    IsPersonanl:boolean;
}

export interface ISharePointFormColumns{
    Name:string;
    Age:any|number;
    Email:string;
    Address:string;
    Compensation:any|number;
    Experience:number;
    Permission:boolean;
    AdminId:any;
    Admin:string|any;
    Manager:any[];
    ManagerId:any[]
}