class RowObject{
    constructor(externalId, objectType, recordName, children, jiraTask, date, initials){
        this.externalId =  externalId;
        this.objectType = objectType; 
        this.recordName = recordName; 
        this.children = children; 
        this.jiraTask = jiraTask; 
        this.date = date;
        this.initials = initials; 
    }
}

module.exports = RowObject;
