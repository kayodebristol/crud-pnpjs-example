import 'es6-promise';
import 'whatwg-fetch';
import { sp } from "@pnp/sp";
import Logger from 'js-logger';
import {config} from './config';

(window).global = window;
if (global === undefined) {
    var global = window;
}

sp.setup({
    sp: {
        ie11: true,
        defaultCachingStore: "local",
        defaultCachingTimeoutSeconds: 360,
        globalCacheDisable: true, // or true to disable caching in case of debugging/testing
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
    console.log(data); 
    return data; 

};

const formatItem = (rawData)=>{
    return {
        Title: `${rawData.title.slice(0, 25)}`,
        POC_x0020_Phone_x0020_No_x002e_: `(${rawData.areaCode}) ${rawData.prefix}-${rawData.number}`
        
        
    }; 
}; 

export const addNewRecord = async (form)=>{
     // Create
    //console.log(form);
    let item = formatItem(form); 
    await sp.web.lists.getByTitle('TroubleTicket').items.add(item)
    

};

export const updateRecord = async(form)=>{
    
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
        

    });
}