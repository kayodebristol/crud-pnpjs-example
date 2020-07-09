import 'es6-promise';
import 'whatwg-fetch';
import { sp } from "@pnp/sp";
import Logger from 'js-logger';
import {config} from './config';

Logger.useDefaults(); 

(window).global = window;
if (global === undefined) {
    var global = window;
}		
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

export const getData = async (config)=>{
    
    const today = new Date().toISOString(); 
        //PnPjs fetch example
    let data = await sp.web.select("Title", "Description").get();
    
    return data; 

};

const formatItem = (rawData)=>{
    return {
        Title: `${rawData.title.slice(0, 25)}`,
        
    }; 
}; 

export const addNewRecord = async (form)=>{
     // Create
    console.log(form);
    let item = formatItem(form); 
    await sp.web.lists.getByTitle('TroubleTicket').items.add(item)
    .then((iar) => {
        let attachments = form.file.attachment.slice(); 
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
        url1 = '//arylether.sharepoint.com/sites/sandbox/Lists/Trouble-Ticket/DispForm.aspx?ID=';
        url2 = '&Source=http%3A%2F%2Farylether%2Esharepoint%2Ecom%2sites%2Esandbox%2ELists%2FTrouble-Ticket%2FAll%2520Items%2Easpx&ContentTypeId=0x01005EC21FC29FB2CD4E851AF576EB8BB710';
        url3 = '//arylether.sharepoint.com/sites/sandbox/Trouble-Ticket';
        
        const emailProps = {
            To: form.user.email !== form.POC.currentSelectedItems[0].tertiaryText && form.user.email !== null ? [form.user.email, form.POC.currentSelectedItems[0].tertiaryText] : [form.POC.currentSelectedItems[0].tertiaryText],
            CC: form.engineer.currentSelectedItems.length > 0 ? [form.engineer.currentSelectedItems[0].tertiaryText] : null,
            Subject: `KMO Support Ticket - ${iar.data.Id} -  ${form.affectedResource.value} ${form.category.value} ${form.description.value.slice(0, 25)} `,
            Body: `<b>form is an auto-generated email.  Please do not respond.</b></br>
                    You have created a KMO Support Ticket. Your ticket number is <b> ${iar.data.Id}</b>.</br>
                    You may view your ticket at the following link: <a href='${url1}${iar.data.Id}${url2}'>
                        ${form.affectedResource.value} ${form.category.value} ${form.description.value.slice(0, 25)} </a>`
        }
        form.saveButton.body = `<p>SUCCESS!</p>`;
        sp.utility.sendEmail(emailProps)
        .then(_ =>{
            setTimeout(()=>{window.open(url3, '_self')},250);
        })
        .catch(e => {
            form.saveButton.body = `<p>Saved! But Email notification failed.</p>`;
            setTimeout(()=>{window.open(url3, '_self')},1000);
        });
        
    
    })
    .catch(e => {
        //form.saveButton.body = `<p>ERROR!</p>`;
    });

    form.type='NEW_FORM';
    if(form.error){form.saveButton.body = `<p>ERROR!</p>`}
};

export const updateRecord = async(form)=>{
    form.type = 'DISPLAY_FORM';
    form.error = false;
    const id = form.id;

    const category = form.category.value !== '' ? form.category.value : null;
    form.category.wasInFocus = true; 
    const affectedResource = form.affectedResource.value !== '' ? form.affectedResource.value : null; 
    form.affectedResource.wasInFocus = true;
    const priority = form.priority.value !== '' ? form.priority.value : null; 
    form.priority.wasInFocus = true; 
    const status = form.status.value !== '' ? form.status.value : null; 
    form.status.wasInFocus = true;
    const description = form.description.value !== '' ? form.description.value : null; 
    form.description.wasInFocus = true; 
    const webApp = form.webApp.value !== '' ? form.webApp.value : null; 
    form.webApp.wasInFocus = true; 
    const navLocation = form.navLocation.value !== '' ? form.navLocation.value : null; 
    form.navLocation.wasInFocus = true; 
    const siteName = form.siteName.value !== '' ? form.siteName.value : null; 
    form.siteName.wasInFocus = true; 

    const siteUrl = form.siteUrl.value !== '' ? form.siteUrl.value : null; 
    form.siteUrl.wasInFocus = true; 
    const POC = form.POC.currentSelectedItems.length > 0 ? form.POC.currentSelectedItems[0].key : null; 
    form.POC.wasInFocus = true;
    const POCPhone = form.POCPhone.value !== '' ? form.POCPhone.value : null; 
    form.POCPhone.wasInFocus = true;
    const section = form.section.value !== '' ? form.section.value : null; 
    form.section.wasInFocus = true;
    const location = form.location.value !== '' ? form.location.value : null; 
    form.location.wasInFocus = true;
    const engineer = form.engineer.currentSelectedItems.length > 0 ? form.engineer.currentSelectedItems[0].key : null; 
    form.engineer.wasInFocus = true;
    const contentManagers = form.contentManagers.currentSelectedItems.length > 0 ? 
    {results:form.contentManagers.currentSelectedItems.map((value, index, array)=>{
        return value.key;
    })} : {results: []}; 
    form.contentManagers.wasInFocus = true;
    const newSiteName = form.newSiteName.value !=='' ? form.newSiteName.value : null;
    form.newSiteName.wasInFocus = true;
    const engineersLog = form.engineersLog.value !== '' ? form.engineersLog.value : null; 
    form.engineersLog.wasInFocus = true;
    const adobeServer = form.adobeServer.value !== '' ? form.adobeServer.value : null; 
    const meetingGroup = form.meetingGroup.value !== '' ? form.meetingGroup.value : null; 
    const meetingName = form.meetingName.value !== '' ? form.meetingName.value : null; 
    const meetingUrl = form.meetingUrl.value !== '' ? form.meetingUrl.value : null; 
    const attachment = form.file.attachment.length > 0  ? form.file.attachment : null; 
    const unit = form.unit.value !== '' ? form.unit.value : null;
    form.unit.wasInFocus = true;
    
    switch (affectedResource) {
        case 'Knowledge Management (SharePoint)':
            if (webApp === null) { 

                form.error=true; 
            } 
            
            if(category === 'New Site'){
                if(newSiteName === null){

                    form.error = true; 
                }
                if (navLocation === null){
                    form.error = true; 
                }
                
                if(contentManagers.results.length < 1){

                form.error = true; 
                }

            }
        break;
        case 'Collaboration (Adobe Connect)':
        default:
            break;
    }

    if (category === null) { 
        form.error = true; 
    } 
    
    if (affectedResource === null) { 
        form.error = true; 
    } 
    
    if (description === null) { 
            form.error = true; 
        } 
    
    if (priority === null) { form.error = true; } 
    
    if (status === null) {  form.error = true; } 
    
    if (POC === null) { form.error = true; } 
    
    if (POCPhone === null) {  form.error = true; } 
    
    if (section === null) { form.error = true; } 
    
    if (location === null) { form.error = true; } 
    
    if (unit === null) { form.error = true; } 
    
    if (!form.error) {
        await sp.web.lists.getByTitle('TroubleTicket').items.getById(id).update({
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
            if (form.file.attachment.length > 0) {
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
            
            url1 = '//arylether.sharepoint.com/TroubleTicket/Lists/TroubleTicket/DispForm.aspx?ID=';
            url2 = '&Source=http%3A%2F%2Farylether%2Esharepoint%2Ecom%2Ftroubleticket%2FLists%2FTroubleTicket%2FAll%2520Items%2Easpx&ContentTypeId=0x01005EC21FC29FB2CD4E851AF576EB8BB710';
            url3 = '//arylethersystems.sharepoint.com/TroubleTicket';
                
            const emailProps = {
                To: form.user.email !== form.POC.currentSelectedItems[0].tertiaryText && form.user.email !== null ? [form.user.email, form.POC.currentSelectedItems[0].tertiaryText] : [form.POC.currentSelectedItems[0].tertiaryText],
                CC: form.engineer.currentSelectedItems.length > 0 ? [form.engineer.currentSelectedItems[0].tertiaryText] : null,
                Subject: `KMO Support Ticket - ${form.id} - ${form.affectedResource.value} ${form.category.value} ${form.description.value.slice(0, 25)} `,
                Body: `<b>form is an auto-generated email.  Please do not respond.</b></br>
                        KMO Support Ticket <b>${form.id}</b> has been updated.</br>
                        
                        Status: <b>${form.status.value}</b></br>
                        POC: <a href='mailto:${form.POC.currentSelectedItems[0].tertiaryText}'>${form.POC.currentSelectedItems[0].primaryText}</a></br>
                        POC Phone: ${form.POCPhone.value}</br>
                        Description: ${form.description.value}</br>
                        Engineer: ${(form.engineer.currentSelectedItems.length > 0 ? "<a href='mailto:" + form.engineer.currentSelectedItems[0].tertiaryText + "'>" + form.engineer.currentSelectedItems[0].primaryText + "</a>" : '')}</br>
                        Engineer's Log: ${form.engineersLog.value} </br>
                        </br>
                        You may view your ticket at the following link: <a href='${url1}${form.id}${url2}'>
                        ${form.affectedResource.value} ${form.category.value} ${form.description.value.slice(0, 25)}</a>`,
            };
            form.updateButton.body = `<p>SUCCESS!</p>`;

            sp.utility.sendEmail(emailProps)
            .then(_ => {
                setTimeout(() => {
                window.open(url3, '_self');
                }, 250);
            })
            .catch(e => {
                form.updateButton.body = `<p>Updated! But Email notification failed.</p>`
                setTimeout(() => {
                    window.open(url3, '_self');
                    }, 1000);
            }); 
        })
        .catch(e => {
            
            form.updateButton.body = `<p>ERROR!</p>`;
        }); 
    }
    form.type = 'EDIT_FORM';
    if(form.error){ form.updateButton.body = `<p>ERROR!</p>`}
}
export const uploadFile = (e) => {
                
    form.attachmentLoading = "loading";
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
        form.file.attachment.push({ name: data.filename, file: data }); 
        
        var url;
        url = '//arylether.sharepoint.com/TroubleTicket/SiteAssets/icons/icons8_';
        

        form.file.items = form.file.attachment.map((value, index, array) => {
            return {
                name: value.file.filename,
                value: value.file.data_uri,
                iconName: `${url}${value.file.filename.slice(-3).toUpperCase()}_32px_1.png`,

            };
        });
        form.attachmentLoading = '';

    });
}