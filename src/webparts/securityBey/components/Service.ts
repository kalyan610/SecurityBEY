import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";

export default class Service {

    public mysitecontext: any;

    public constructor(siteUrl: string, Sitecontext: any) {
        this.mysitecontext = Sitecontext;

        sp.setup({
            sp: {
                baseUrl: siteUrl

            },
        });

    }
    public async isCurrentUserMemberOfGroup1(groupName: string) {
        return await sp.web.currentUser.groups().then((groups: any) => {
            let groupExist = false;
            groups.map((group: any) => {
                if (group.Title = groupName) {
                    groupExist = true;
                }
            });
            return groupExist;
        });

    }

    public async ClerenceList(): Promise<any> {

        return await sp.web.lists.getByTitle("CLEARANCELIST").items.select('Title', 'ID', 'AccessCLEARANCELIST').expand().get().then(function (data) {

            return data;

        });

    }
    public async GetListNameandURL(): Promise<any> {

        return await sp.web.lists.getByTitle("URLandListname").items.select('Title', 'URL', 'ColName').expand().get().then(function (data) {

            return data;

        });

    }
    public async getUserByLogin(LoginName: string): Promise<any> {

        try {

            const user = await sp.web.siteUsers.getByLoginName(LoginName).get();

            return user;

        } catch (error) {

            console.log(error);

        }

    }

    public async GetProjectandLocation(): Promise<any> {

        return await sp.web.lists.getByTitle("ProjectandLocation").items.select('Title', 'ID').expand().get().then(function (data) {

            return data;

        });

    }
    public async getCurrentUserSiteGroups(): Promise<any[]> {

        try {

            return (await sp.web.currentUser.groups.select("Id,Title,Description,OwnerTitle,OnlyAllowMembersViewMembership,AllowMembersEditMembership,Owner/Id,Owner/LoginName").expand('Owner').get());

        }
        catch {
            throw 'get current user site groups failed.';
        }

    }

    public async getItemByID(ItemID: any): Promise<any> {
        try {
            const selectedList = 'SecuredBayAccess';
            const Item: any[] = await sp.web.lists.getByTitle(selectedList).items
                .select("*,Title,CapcoEmployeeCode,RequestorName/EMail,EMPSIGN/EMail,ApproverSign/EMail,ApproverNames/EMail,Location/Id,Location/Title")
                .expand("RequestorName,EMPSIGN,ApproverSign,ApproverNames,Location")
                .filter("ID eq '" + ItemID + "'")
                .get();
            return Item[0];
        } catch (error) {
            console.log(error);
        }
    }
    public async getUserByEmail(LoginName: string): Promise<any> {
        try {
            const user = await sp.web.siteUsers.getByEmail(LoginName).get();
            return user;
        } catch (error) {
            console.log(error);
        }
    }


    public async updateEmp(MyRecordId: number) {

        let Myval = 'Completed';

        let MyListTitle = 'SecuredBayAccess';

        try {

            let Filemal = [];

            let list = sp.web.lists.getByTitle(MyListTitle);
            let Varmyval = await list.items.getById(MyRecordId).update({

                //Emp Update



                Title: 'Test1'





            }).then(async r => {

                return Myval;

            })

            return Varmyval;

        }


        catch (error) {
            console.log(error);
        }

    }

    public async getCurrentUser(): Promise<any> {

        try {

            return await sp.web.currentUser.get().then(result => {

                return result;

            });

        } catch (error) {

            console.log(error);

        }

      }



















}
