import { values } from "office-ui-fabric-react"
import { resultContent } from "office-ui-fabric-react/lib/components/FloatingPicker/PeoplePicker/PeoplePicker.scss";
import { Web } from "sp-pnp-js"

export default class formservices {
    constructor() {

    }

    public removeUserFromGroup = (web: Web, groupName: string, UserName: string[]): Promise<any> => {
        return new Promise(async (resolve, reject) => {
            let countSuccess = 0;
            for (var user of UserName) {
                await web.siteGroups.getByName(groupName).users.removeByLoginName(user)
                    .then(value => {
                        countSuccess++;
                    }).catch(err => {
                        console.log(err);
                    });
            }
            countSuccess == UserName.length ? resolve(true) : resolve(false);
        })
    }

    public addUserToGroup = (web: Web, groupName: string, user: any): Promise<any> => {
        return new Promise((resolve, reject) => {
            web.siteGroups.getByName(groupName).users.add(user.UserName)
            .then(data => {
                resolve(true);
              }).catch(err => {
                resolve(false);
              })
        })
    }
}