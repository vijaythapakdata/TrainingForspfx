import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ListNames } from "../../enum/ListNames";

export default class ChoiceFieldsClass {
    private context: WebPartContext;
    constructor(context: WebPartContext) {
        this.context = context;
    }

    public async getChoiceValues(siteurl: string, fieldvalue: string): Promise<any> { //My Name=MyName

        try {
            const response = await fetch(`${siteurl}/_api/web/lists/getbytitle('${ListNames.FirstList}')/fields?$filter=EntityPropertyName eq '${fieldvalue}'`,

                {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json;odata=nometadata'
                    }
                }
            );

            if (!response.ok) {
                throw new Error(`error while fetching the choice values ${response.text}-${response.statusText}`);
            }

            const data = await response.json();
            const choice = data.value[0].Choices;
            return choice.map((items: any) => ({
                key: items,
                text: items
            }));
        }
        catch (err) {
            console.error(err);
            return [];

        }
    }
    //get lookup values

    public async getLookupValues(): Promise<any> {
        try {
            const response = await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${ListNames.Lookup}')/items?$select=Title,ID`,
                {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json;odata=nometadata'
                    }
                }
            );
            if (!response.ok) {
                throw new Error(`error while fetching the lookup values ${response.text}-${response.statusText}`);
            }

            const data = await response.json();
            return data.value.map((city: { Title: string, ID: string }) => ({
                key: city.ID,
                text: city.Title
            }));
        }

        catch (err) {
            console.error(err);
            return []
        }
    }

}