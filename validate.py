
import counts


def isMacUnique(mac_to_check, staging_mac_col, staging_df):

    filt = staging_df.iloc[:, staging_mac_col] == mac_to_check
    rowData = staging_df[filt]

    if len(rowData) == 0:
        print(f'MAC {mac_to_check} was not found in Staging Sheet, skipping renaming')
        counts.mac_not_found += 1
        return False

    if len(rowData) == 1:
        return True
    
    print(f'MAC {mac_to_check} was duplicated on Staging, skipping renaming')
    counts.mac_duplicated += 1
    return False


def isSerialUnique(serial_to_check, serial_staging_col, staging_df):
    filt = staging_df.iloc[:, serial_staging_col] == serial_to_check
    rowData = staging_df[filt]

    if len(rowData) == 0:
        print(f'Serial {serial_to_check} was not found in Staging Sheet, skipping renaming')
        counts.serial_not_found += 1
        return False

    if len(rowData) == 1:
        return True
    
    print(f'Serial {serial_to_check} was duplicated on Staging Sheet, skipping renaming')
    counts.serial_duplicated += 1
    return False


def isMacAndSerialOnSameRow(export_mac, export_serial, staging_mac_col, staging_serial_col, staging_df):

    macFilter = staging_df.iloc[:, staging_mac_col] == export_mac
    stagingRow = staging_df[macFilter]

    if len(stagingRow) == 1:
        stagingSerial = stagingRow.iloc[0][staging_serial_col]

        if stagingSerial == export_serial:
            return True
        
        print(f'MAC {export_mac} does not have serial {export_serial} on Staging Sheet, double check scanning')
        counts.mac_serial_mismatch += 1
        return False

    return False
    

def isNotAlreadyRenamed(export_row, export_mac, export_name_col):

    if export_row[export_name_col] == export_mac.lower():
        return True
    
    print(f'Device with MAC {export_mac} on export sheet has already been renamed, skipping renaming')
    counts.already_renamed += 1
    return False
