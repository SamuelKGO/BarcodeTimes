#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include "BluetoothSerial.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>

#define MAX_BARCODES 1000

struct barcode_data {
    char barcode[256];
    time_t start_time;
    time_t end_time;
};

void export_to_excel(struct barcode_data *data, int num_barcodes) {
    char file_name[256];
    time_t now = time(NULL);
    struct tm *tm = localtime(&now);
    strftime(file_name, sizeof(file_name), "TS_data_%m-%d-%Y--%I-%M-%S%p.xlsx", tm);

    FILE *fp = fopen(file_name, "w");
    if (fp == NULL) {
        printf("Error opening file %s\n", file_name);
        return;
    }

    fprintf(fp, "Barcode\tStart Time\tEnd Time\tDuration (minutes)\n");

    for (int i = 0; i < num_barcodes; i++) {
        struct barcode_data *bd = &data[i];
        double duration = difftime(bd->end_time, bd->start_time) / 60.0;
        fprintf(fp, "%s\t%s\t%s\t%.2f\n", bd->barcode, ctime(&bd->start_time), ctime(&bd->end_time), duration);
    }

    fclose(fp);
}

int main() {
    struct barcode_data barcodes[MAX_BARCODES];
    int num_barcodes = 0;

    while (1) {
        char datetime[256];
        char barcode[256];
        printf("Scan barcode: ");
        scanf("%s", barcode);

        if (strlen(barcode) == 0) {
            break;
        }

        if (strcmp(barcode, "q") == 0 || strcmp(barcode, "Q") == 0) {
            break;
        }

        int found = 0;
        for (int i = 0; i < num_barcodes; i++) {
            if (strcmp(barcodes[i].barcode, barcode) == 0) {
                barcodes[i].end_time = time(NULL);
                found = 1;
                break;
            }
        }

        if (!found) {
            struct barcode_data *bd = &barcodes[num_barcodes++];
            strcpy(bd->barcode, barcode);
            bd->start_time = time(NULL);
            bd->end_time = bd->start_time;
        }
    }

    export_to_excel(barcodes, num_barcodes);

    return 0;
}
