import createReport from "docx-templates";

const templateFilePath = 'sample.docx';

export declare interface WorkData {
    startDate: string;
    workName: string;
    endDate: string;
    result: string;
}

class DocumentGanaration {

    template!: Buffer;
    buffer!: BlobPart;
    workList: Array<WorkData>;
    private intervalLoadTemplate: number | undefined;
    private intervalLoadBuffer: number | undefined;

    private async saveWithSpecialName() {
        let reportName = "test_sample.docx";
        if (this.buffer) {
            await this.saveFile(reportName);
        }
        else {
            this.intervalLoadBuffer = window.setInterval(async () => {
                if (this.buffer) {
                    clearInterval(this.intervalLoadBuffer);
                    await this.saveFile(reportName);
                }
            }, 10);
        }
    }

    constructor() {
        this.generateDocument();
        this.workList = [
            {
                startDate: "01.03",
                workName: "some work1",
                endDate: "02.04",
                result: "secsses"
            },
            {
                startDate: "01.03",
                workName: "some work2",
                endDate: "02.04",
                result: "secsses"
            },
            {
                startDate: "01.03",
                workName: "some work3",
                endDate: "02.04",
                result: "secsses"
            },
            {
                startDate: "01.03",
                workName: "some work4",
                endDate: "02.04",
                result: "secsses"
            }
        ]
    }

    private generateDocument(): void {
        fetch(templateFilePath)
            .then((response) => {
                response.arrayBuffer().then((arrayBuffer) => {
                    this.template = Buffer.alloc(arrayBuffer.byteLength)
                    Buffer.from(arrayBuffer).copy(this.template);
                })
            });

    }

    async writeData() {
        clearInterval(this.intervalLoadTemplate);
        if (!this.checkDataAndWrite()) {
            this.intervalLoadTemplate = window.setInterval(async () => {
                this.checkDataAndWrite();
            }, 10);
        }

    }

    //write date in template
    private async writeDataWithTemplate() {

        this.buffer = await createReport({
            template: this.template,
            data: {
                workList: this.workList
            },
        });
        return this.buffer;
    }

    private checkDataAndWrite() {
        if (this.template) {
            clearInterval(this.intervalLoadTemplate);
            this.writeDataWithTemplate().then(() => {
                this.saveWithSpecialName();
            });
        }
        return this.template;
    }

    private async saveFile(fileName: string) {
        await this.downloadFile(this.buffer, fileName, "application/msword");
    }

    async downloadFile(data: BlobPart, filename: string, mime: string) {
        // It is necessary to create a new blob object with mime-type explicitly set
        // otherwise only Chrome works like it should
        const blob = new Blob([data], { type: mime || 'application/octet-stream' });
        if (typeof window.navigator.msSaveBlob !== 'undefined') {
            // IE doesn't allow using a blob object directly as link href.
            // Workaround for "HTML7007: One or more blob URLs were
            // revoked by closing the blob for which they were created.
            // These URLs will no longer resolve as the data backing
            // the URL has been freed."

            return window.navigator.msSaveBlob(blob, filename);
        }
        // Other browsers
        // Create a link pointing to the ObjectURL containing the blob
        const blobURL = window.URL.createObjectURL(blob);
        const tempLink = document.createElement('a');
        tempLink.style.display = 'none';
        tempLink.href = blobURL;
        tempLink.setAttribute('download', filename);
        // Safari thinks _blank anchor are pop ups. We only want to set _blank
        // target if the browser does not support the HTML5 download attribute.
        // This allows you to download files in desktop safari if pop up blocking
        // is enabled.
        if (typeof tempLink.download === 'undefined') {
            tempLink.setAttribute('target', '_blank');
        }
        document.body.appendChild(tempLink);
        tempLink.click();
        document.body.removeChild(tempLink);
        window.URL.revokeObjectURL(blobURL);
        setTimeout(() => {
            // For Firefox it is necessary to delay revoking the ObjectURL

        }, 10);
        //await new Promise(r => setTimeout(r, 100));
    }
}

export default DocumentGanaration;