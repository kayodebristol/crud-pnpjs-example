import 'es6-promise';
import 'whatwg-fetch';
import { sp } from "@pnp/sp";
import Logger from 'js-logger';
    
Logger.useDefaults(); 

(window).global = window;
if (global === undefined) {
    var global = window;
}		

export let getData = async (config)=>{
    Logger.debug(config)
    const today = new Date().toISOString(); 
    sp.setup({
        sp: {
            ie11: true,
            defaultCachingStore: "local", // or "local"
            defaultCachingTimeoutSeconds: 360,
            globalCacheDisable: false, // or true to disable caching in case of debugging/testing
            headers: {
                Accept: "application/json;odata=verbose",
            },
            baseUrl: config.baseUrl
        }
    });
    
    //PnPjs fetch example
    let data = await sp.web.select("Title", "Description").get()
    Logger.debug(data); 
    ///return data; 

   


}

export let addNewRecord = async ()=>{
     // Create

    await sp.web.lists.getByTitle('Trouble-Ticket').items.add({
        Title: `${affectedResource} ${category} ${description.slice(0, 25)}`,
        Category: category,
        Affected_x0020_Resource: affectedResource,
        Priority: priority,
        Status: status,
        Top_x0020_Level_x0020_Site: webApp,
        Navigation_x0020_Element: navLocation,
        Site_x0020_Name: siteName,
        Site_x0020_URL: siteUrl,
        POCId: POC,
        EngineerId: engineer,
        Section: section,
        Unit: unit,
        Location: location,
        Engineers_x0020_Log: engineersLog,
        POC_x0020_Phone: POCPhone,
        Description: description,
        New_x0020_Site_x0020_Name: newSiteName, 
        Content_x0020_ManagersId: contentManagers,

    })
    .then((iar) => {
        let attachments = this.file.attachment.slice(); 
        if (attachments.length > 0) {
            for (var j = 0; j < attachments.length; j++) {
                var base64Marker = ";base64,"; var base64Index = attachments[j].file.data_uri.indexOf(base64Marker) + base64Marker.length; var base64 = attachments[j].file.data_uri.substring(base64Index);
                var raw = window.atob(base64);
                var rawLength = raw.length;
                var array = new window.Uint8Array(new window.ArrayBuffer(rawLength));

                for (var i = 0; i < rawLength; i++) {
                    array[i] = raw.charCodeAt(i);
                }

                iar.item.attachmentFiles.add(attachments[j].name, array);
            }

        }
        return iar;
    })
    .then((iar)=>{
        var url1, url2, url3;
        switch(network){
            case 'NIPR': 
                url1 = '//portal.usafricom.mil/sites/troubleticket/Lists/TroubleTicket/DispForm.aspx?ID=';
                url2 = '&Source=http%3A%2F%2Fportal%2Eafghan%2Eswa%2Earmy%2Emil%2Fsites%2Ftroubleticket%2FLists%2FTroubleTicket%2FAll%2520Items%2Easpx&ContentTypeId=0x01005EC21FC29FB2CD4E851AF576EB8BB710';
                url3 = '//portal.usafricom.mil/support';
                break;
            case 'SIPR': 
                url1 = '//portal.usafricom.smil.mil/sites/troubleticket/Lists/TroubleTicket/DispForm.aspx?ID=';
                url2 = '&Source=http%3A%2F%2Fportal%2Eafghan%2Eswa%2Earmy%2Esmil%2Emil%2Fsites%2Ftroubleticket%2FLists%2FTroubleTicket%2FAll%2520Items%2Easpx&ContentTypeId=0x01005EC21FC29FB2CD4E851AF576EB8BB710';
                url3 = '//portal.usafricom.smil.mil/support';
                break;
            default: 
            
                url1 = '//arylether.sharepoint.com/sites/sandbox/Lists/Trouble-Ticket/DispForm.aspx?ID=';
                url2 = '&Source=http%3A%2F%2Farylether%2Esharepoint%2Ecom%2sites%2Esandbox%2ELists%2FTrouble-Ticket%2FAll%2520Items%2Easpx&ContentTypeId=0x01005EC21FC29FB2CD4E851AF576EB8BB710';
                url3 = '//arylether.sharepoint.com/sites/sandbox/Trouble-Ticket';

                break;
        }

        const emailProps = {
            To: this.user.email !== this.POC.currentSelectedItems[0].tertiaryText && this.user.email !== null ? [this.user.email, this.POC.currentSelectedItems[0].tertiaryText] : [this.POC.currentSelectedItems[0].tertiaryText],
            CC: this.engineer.currentSelectedItems.length > 0 ? [this.engineer.currentSelectedItems[0].tertiaryText] : null,
            Subject: `KMO Support Ticket - ${iar.data.Id} -  ${this.affectedResource.value} ${this.category.value} ${this.description.value.slice(0, 25)} `,
            Body: `<b>This is an auto-generated email.  Please do not respond.</b></br>
                    You have created a KMO Support Ticket. Your ticket number is <b> ${iar.data.Id}</b>.</br>
                    You may view your ticket at the following link: <a href='${url1}${iar.data.Id}${url2}'>
                        ${this.affectedResource.value} ${this.category.value} ${this.description.value.slice(0, 25)} </a>`
        }
        this.saveButton.body = <p>SUCCESS!</p>
        sp.utility.sendEmail(emailProps)
        .then(_ =>{
            setTimeout(()=>{window.open(url3, '_self')},250);
        })
        .catch(e => {
            this.saveButton.body = <p>Saved! But Email notification failed.</p>;
            setTimeout(()=>{window.open(url3, '_self')},1000);
        });
        
    
    })
    .catch(e => {
        this.saveButton.body = <p>ERROR!</p>;
    });

    this.type='NEW_FORM';
    if(this.error){this.saveButton.body = <p>ERROR!</p>}
}

export let updateRecord = async(){
    this.type = 'DISPLAY_FORM'
        this.error = false;
        const id = this.id;

        const category = this.category.value !== '' ? this.category.value : null;
        this.category.wasInFocus = true; 
        const affectedResource = this.affectedResource.value !== '' ? this.affectedResource.value : null; 
        this.affectedResource.wasInFocus = true;
        const priority = this.priority.value !== '' ? this.priority.value : null; 
        this.priority.wasInFocus = true; 
        const status = this.status.value !== '' ? this.status.value : null; 
        this.status.wasInFocus = true;
        const description = this.description.value !== '' ? this.description.value : null; 
        this.description.wasInFocus = true; 
        const webApp = this.webApp.value !== '' ? this.webApp.value : null; 
        this.webApp.wasInFocus = true; 
        const navLocation = this.navLocation.value !== '' ? this.navLocation.value : null; 
        this.navLocation.wasInFocus = true; 
        const siteName = this.siteName.value !== '' ? this.siteName.value : null; 
        this.siteName.wasInFocus = true; 

        const siteUrl = this.siteUrl.value !== '' ? this.siteUrl.value : null; 
        this.siteUrl.wasInFocus = true; 
        const POC = this.POC.currentSelectedItems.length > 0 ? this.POC.currentSelectedItems[0].key : null; 
        this.POC.wasInFocus = true;
        const POCPhone = this.POCPhone.value !== '' ? this.POCPhone.value : null; 
        this.POCPhone.wasInFocus = true;
        const section = this.section.value !== '' ? this.section.value : null; 
        this.section.wasInFocus = true;
        const location = this.location.value !== '' ? this.location.value : null; 
        this.location.wasInFocus = true;
        const engineer = this.engineer.currentSelectedItems.length > 0 ? this.engineer.currentSelectedItems[0].key : null; 
        this.engineer.wasInFocus = true;
        const contentManagers = this.contentManagers.currentSelectedItems.length > 0 ? 
        {results:this.contentManagers.currentSelectedItems.map((value, index, array)=>{
            return value.key;
        })} : {results: []}; 
        this.contentManagers.wasInFocus = true;
        const newSiteName = this.newSiteName.value !=='' ? this.newSiteName.value : null;
        this.newSiteName.wasInFocus = true;
        const engineersLog = this.engineersLog.value !== '' ? this.engineersLog.value : null; 
        this.engineersLog.wasInFocus = true;
        const adobeServer = this.adobeServer.value !== '' ? this.adobeServer.value : null; 
        const meetingGroup = this.meetingGroup.value !== '' ? this.meetingGroup.value : null; 
        const meetingName = this.meetingName.value !== '' ? this.meetingName.value : null; 
        const meetingUrl = this.meetingUrl.value !== '' ? this.meetingUrl.value : null; 
        const attachment = this.file.attachment.length > 0  ? this.file.attachment : null; 
        const unit = this.unit.value !== '' ? this.unit.value : null;
        this.unit.wasInFocus = true;
        
        switch (affectedResource) {
            case 'Knowledge Management (SharePoint)':
                if (webApp === null) { 

                    this.error=true; 
                } 
                
               if(category === 'New Site'){
                   if(newSiteName === null){

                       this.error = true; 
                   }
                   if (navLocation === null){
                       this.error = true; 
                   }
                   
                   if(contentManagers.results.length < 1){

                    this.error = true; 
                   }

               }
            break;
            case 'Collaboration (Adobe Connect)':
            default:
                break;
        }

        if (category === null) { 
            this.error = true; 
        } 
        
        if (affectedResource === null) { 
            this.error = true; 
        } 
        
        if (description === null) { 
             this.error = true; 
            } 
        
        if (priority === null) { this.error = true; } 
        
        if (status === null) {  this.error = true; } 
        
        if (POC === null) { this.error = true; } 
        
        if (POCPhone === null) {  this.error = true; } 
        
        if (section === null) { this.error = true; } 
        
        if (location === null) { this.error = true; } 
        
        if (unit === null) { this.error = true; } 
        
        if (!this.error) {
            await sp.web.lists.getByTitle('Trouble-Ticket').items.getById(id).update({
                Id: id,
                Title: `${affectedResource} ${category} ${description.slice(0, 25)}`,
                Category: category,
                Affected_x0020_Resource: affectedResource,
                Priority: priority,
                Status: status,
                Top_x0020_Level_x0020_Site: webApp,
                Navigation_x0020_Element: navLocation,
                Site_x0020_Name: siteName,
                Site_x0020_URL: siteUrl,
                POCId: POC,
                EngineerId: engineer,
                Section: section,
                Unit: unit,
                Location: location,
                Engineers_x0020_Log: engineersLog,
                POC_x0020_Phone: POCPhone,
                Description: description,
                New_x0020_Site_x0020_Name: newSiteName, 
                Content_x0020_ManagersId: contentManagers,


            })
            .then((iar) => {
                if (this.file.attachment.length > 0) {
                    var base64Marker = ";base64,";
                    var base64Index = attachment.file.data_uri.indexOf(base64Marker) + base64Marker.length; var base64 = attachment.file.data_uri.substring(base64Index);
                    var raw = window.atob(base64);
                    var rawLength = raw.length;
                    var array = new window.Uint8Array(new window.ArrayBuffer(rawLength));

                    for (var i = 0; i < rawLength; i++) {
                        array[i] = raw.charCodeAt(i);
                    }
                    iar.item.attachmentFiles.add(attachment.file.filename, array);
                } 
                var url1, url2, url3;
                switch (process.env.REACT_APP_ENV) {
                    case 'NIPR':
                        url1 = '//portal.afghan.swa.army.mil/sites/troubleticket/Lists/TroubleTicket/DispForm.aspx?ID=';
                        url2 = '&Source=http%3A%2F%2Fportal%2Eafghan%2Eswa%2Earmy%2Emil%2Fsites%2Ftroubleticket%2FLists%2FTroubleTicket%2FAll%2520Items%2Easpx&ContentTypeId=0x01005EC21FC29FB2CD4E851AF576EB8BB710';
                        url3 = '//portal.afghan.swa.army.mil/support';
                        break;
                    case 'SIPR':
                        url1 = '//portal.afghan.swa.army.smil.mil/sites/troubleticket/Lists/TroubleTicket/DispForm.aspx?ID=';
                        url2 = '&Source=http%3A%2F%2Fportal%2Eafghan%2Eswa%2Earmy%2Esmil%2Emil%2Fsites%2Ftroubleticket%2FLists%2FTroubleTicket%2FAll%2520Items%2Easpx&ContentTypeId=0x01005EC21FC29FB2CD4E851AF576EB8BB710';
                        url3 = '//portal.afghan.swa.army.smil.mil/support';
                        break;
                    case 'CXSWA':
                        url1 = '//portal.afgn.centcom.gctf.cmil.mil/sites/troubleticket/Lists/TroubleTicket/DispForm.aspx?ID=';
                        url2 = '&Source=http%3A%2F%2Fportal%2Eafgn%2Ecentcom%2Egctf%2Ecmil%2Emil%2Fsites%2Ftroubleticket%2FLists%2FTroubleTicket%2FAll%2520Items%2Easpx&ContentTypeId=0x01005EC21FC29FB2CD4E851AF576EB8BB710';
                        url3 = '//portal.afgn.centcom.gctf.cmil.mil/support';
                        break;
                    default:
                        url1 = '//arylethersystems.sharepoint.com/TroubleTicket/Lists/TroubleTicket/DispForm.aspx?ID=';
                        url2 = '&Source=http%3A%2F%2Farylethersystems%2Esharepoint%2Ecom%2Ftroubleticket%2FLists%2FTroubleTicket%2FAll%2520Items%2Easpx&ContentTypeId=0x01005EC21FC29FB2CD4E851AF576EB8BB710';
                        url3 = '//arylethersystems.sharepoint.com/TroubleTicket';
                        break;
                }
                    
                const emailProps = {
                    To: this.user.email !== this.POC.currentSelectedItems[0].tertiaryText && this.user.email !== null ? [this.user.email, this.POC.currentSelectedItems[0].tertiaryText] : [this.POC.currentSelectedItems[0].tertiaryText],
                    CC: this.engineer.currentSelectedItems.length > 0 ? [this.engineer.currentSelectedItems[0].tertiaryText] : null,
                    Subject: `KMO Support Ticket - ${this.id} - ${this.affectedResource.value} ${this.category.value} ${this.description.value.slice(0, 25)} `,
                    Body: `<b>This is an auto-generated email.  Please do not respond.</b></br>
                            KMO Support Ticket <b>${this.id}</b> has been updated.</br>
                            
                            Status: <b>${this.status.value}</b></br>
                            POC: <a href='mailto:${this.POC.currentSelectedItems[0].tertiaryText}'>${this.POC.currentSelectedItems[0].primaryText}</a></br>
                            POC Phone: ${this.POCPhone.value}</br>
                            Description: ${this.description.value}</br>
                            Engineer: ${(this.engineer.currentSelectedItems.length > 0 ? "<a href='mailto:" + this.engineer.currentSelectedItems[0].tertiaryText + "'>" + this.engineer.currentSelectedItems[0].primaryText + "</a>" : '')}</br>
                            Engineer's Log: ${this.engineersLog.value} </br>
                            </br>
                            You may view your ticket at the following link: <a href='${url1}${this.id}${url2}'>
                            ${this.affectedResource.value} ${this.category.value} ${this.description.value.slice(0, 25)}</a>`,
                };
                this.updateButton.body = <p>SUCCESS!</p>

                sp.utility.sendEmail(emailProps)
                .then(_ => {
                    setTimeout(() => {
                    window.open(url3, '_self');
                    }, 250);
                })
                .catch(e => {
                    this.updateButton.body = <p>Updated! But Email notification failed.</p>
                    setTimeout(() => {
                        window.open(url3, '_self');
                        }, 1000);
                }); 
            })
            .catch(e => {
                
                this.updateButton.body = <p>ERROR!</p>
            }); 
        }
        this.type = 'EDIT_FORM';
        if(this.error){ this.updateButton.body = <p>ERROR!</p>}
    }
    uploadFile(e) {
                
        this.attachmentLoading = "loading";
        //let attachment = [];
        //const reader = new FileReader();
        const file = e.target.files[0];
        e.target.value = null;
        return new Promise((resolve, reject) => {
            var fr = new FileReader(); fr.onload = (upload) => {
                resolve({
                    data_uri: upload.target.result,
                    filename: file.name,
                    filetype: file.type
                });
            };
            fr.readAsDataURL(file)
        })
            .then(data => {
                this.file.attachment.push({ name: data.filename, file: data }); 
                
                var url;
                switch (process.env.REACT_APP_ENV) {
                    case 'NIPR':
                        url = '//portal.afghan.swa.army.mil/sites/troubleticket/SiteAssets/icons/icons8_';
                        break;
                    case 'SIPR':
                        url = '//portal.afghan.swa.army.smil.mil/sites/troubleticket/SiteAssets/icons/icons8_';
                        break;
                    case 'CXSWA':
                        url = '//portal.afgn.centcom.gctf.cmil.mil/sites/troubleticket/SiteAssets/icons/icons8_';
                        break;
                    default:
                        url = '//arylethersystems.sharepoint.com/TroubleTicket/SiteAssets/icons/icons8_';
                        break;
                }

                this.file.items = this.file.attachment.map((value, index, array) => {
                    return {
                        name: value.file.filename,
                        value: value.file.data_uri,
                        iconName: `${url}${value.file.filename.slice(-3).toUpperCase()}_32px_1.png`,

                    };
                });
                this.attachmentLoading = '';

            });
}