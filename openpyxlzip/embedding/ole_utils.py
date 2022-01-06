# import olefile
import binascii
import math
import os

# https://olefile.readthedocs.io/en/latest/Howto.html
# https://github.com/decalage2/olefile/issues/6
# https://github.com/decalage2/olefile/blob/master/olefile/olefile.py
# https://olefile.readthedocs.io/en/latest/olefile.html
# https://pypi.org/project/pyemf/
# http://pyemf.sourceforge.net/
# https://www.openoffice.org/sc/compdocfileformat.pdf
# file:///home/matthew/ecsite/temp_files/excel_test_files/[MS-CFB].pdf

class OLEFile:
    DIFAT_SECT=-4
    FAT_SECT=-3
    ENDOFCHAIN=-2
    FREESECT=-1

    def __init__(self, bin_filename, display_filename, add_payload_prefix=True):
        self.display_filename = display_filename
        self.add_payload_prefix = add_payload_prefix
        self.long_filename = "C:\\Users\\12149\\Downloads\\rdfc\\{}".format(self.display_filename)
        self.temp_filename = "C:\\Users\\12149\\AppData\\Local\\Temp\\{{1AC28F20-7FA0-4EC6-B88E-C4E6BBD2A693}}\\{}".format(self.display_filename)
        self.byteorder = "little"    #If little reverse all byte arrays, The 32-bit integer value 13579BDFH is converted into the Little-Endian byte sequence DFH 9BH 57H 13H,
        self.sector_size = 512
        self.sec_ids_per_msat = int(self.sector_size / 4) # 128
        self.header_chain_entries = 109
        self.bin_filename = bin_filename
        self.short_sec_size = 64
        self.header_msat = []
        self.all_msats = []
        self.sat = []
        self.header_fields = {
            "magic": "D0 CF 11 E0 A1 B1 1A E1", #Magic
            "uuid": " ".join(["00"] * 16),    #UUID   must be zeros
            "rev_num": "3E 00",  #Rev number - must be 3E
            "version_num": "03 00",  #Version num - can be 03 or 04
            "endianness": "FE FF",  #FF FE - big endian, FE FF - must be little endian
            "sec_size": "09 00",  #sec size must be 512
            "ssec_size": "06 00",  #ssec size must be 64
            "reserved": "00 00 00 00 00 00",      #Reserved  - must be zeros
            "num_dir_sectors": 0,      #Num dir sectors  - must be zeros
            # "num_fat_sectors": "02 00 00 00",  #Num FAT sectors Total number of sectors used for the sector allocation table (➜5.2) variable
            "first_sec_id": 1, #SecID of first sector of the directory stream (➜7)   seems to always be 1, this points to 0400 for the root entry
            "not_used": "00 00 00 00",   #Not used - must be zeros
            "cutoff_size": "00 10 00 00",   #Mini Stream Cutoff Size - must be 00 10 00 00 or 4096
            "first_ssat_secid": 2,    #SecID of first sector of the short-sector allocation table (➜6.2), or –2 (End Of Chain SecID, ➜3.1) if not extant
            "total_secs_ssat": 1,    #Total number of sectors used for the short-sector allocation table (➜6.2)
            "first_msat_secid": -2,    #SecID of first sector of the master sector allocation table (➜5.1), or –2 (End Of Chain SecID, ➜3.1) if no additional sectors used
            "total_secs_msat": 0,    #Total number of sectors used for the master sector allocation table (➜5.1)
        }
        self.root_dir = {
            "name": "52 00 6F 00 6F 00 74 00 20 00 45 00 6E 00 74 00 72 00 79 00",
            "name_len": "16 00",
            "type": "05",
            "color": "00", #black
            "left": "FF FF FF FF",
            "right": "FF FF FF FF",
            "child": "02 00 00 00",
            "uuid": "65 CA 01 B8 FC A1 D0 11 85 AD 44 45 53 54 00 00",
            "starting_sector": 3,
            "stream_size": 192
        }
        self.comp_obj_dir = {
            "name": "01 00 43 00 6F 00 6D 00 70 00 4F 00 62 00 6A 00",
            "name_len": "12 00",
            "type": "02",
            "color": "01", #red
            "left": "01 00 00 00",
            "right": "03 00 00 00",
            "child": "FF FF FF FF",
            "starting_sector": 1,
            "stream_size": 93
        }
        self.ole_dir = {
            "name": "01 00 4F 00 6C 00 65 00",
            "name_len": "0A 00",
            "type": "02",
            "color": "00", #black
            "left": "FF FF FF FF",
            "right": "FF FF FF FF",
            "child": "FF FF FF FF",
            "starting_sector": 0,
            "stream_size": 20
        }
        self.ole_native_dir = {
            "name": "01 00 4F 00 6C 00 65 00 31 00 30 00 4e 00 61 00 74 00 69 00 76 00 65 00",
            "name_len": "1A 00",
            "type": "02",
            "color": "00", #black
            "left": "FF FF FF FF",
            "right": "FF FF FF FF",
            "child": "FF FF FF FF",
            "starting_sector": 5,   #Variable    
        }
        self.obj_prefix = {
        }
        self.empty_dir = {
            "name": "00",
            "name_len": "00 00",
            "type": "00",
            "color": "00", #black
            "left": "FF FF FF FF",
            "right": "FF FF FF FF",
            "child": "FF FF FF FF",
            "starting_sector": 0,
            "stream_size": 0
        }
        self.contents_dir = {
            "name": "43 00 4F 00 4E 00 54 00 45 00 4E 00 54 00 53 00",
            "name_len": "12 00",
            "type": "02",
            "color": "00", #black
            "left": "FF FF FF FF",
            "right": "FF FF FF FF",
            "child": "FF FF FF FF",
            "starting_sector": 5, #This is 
            # "stream_size": "CB 98 01 00 00 00 00 00"
        }

    def build_sat_and_msat(self):
        self.num_sectors = math.ceil(self.payload_len / self.sector_size) #Correct
        self.num_short_sectors = math.ceil(self.payload_len / 64) + 2
        self.initial_offset = 4
        if self.payload_len < 1024:
            self.initial_offset = 2
        self.ssat_sectors_for_payload = math.ceil((self.num_sectors - self.header_chain_entries) / self.sec_ids_per_msat) #Original
        self.total_sec_ids = math.ceil((self.num_sectors - (self.header_chain_entries - 1)) / self.sec_ids_per_msat)
        self.total_secs_msat = 0
        self.start_sec_id = self.initial_offset
        if self.payload_len >= 1024:
            self.build_sat_and_msat_long_file()
        else:
            self.build_sat_and_msat_short_file()
        print(self.total_secs_msat, self.total_sec_ids, self.header_chain_entries, self.sec_ids_per_msat)
        print("payload len", self.payload_len, self.num_sectors, self.total_sec_ids, self.ssat_sectors_for_payload, self.contents_starting_sector, self.total_secs_msat)

    def build_sat_and_msat_short_file(self):
        #Fill in dir fields
        self.num_fat_sectors = math.ceil((self.num_sectors) / (self.sec_ids_per_msat - 1))
        self.contents_starting_sector = self.initial_offset + self.ssat_sectors_for_payload
        self.root_dir["stream_size"] = int((self.num_short_sectors) * 64)
        self.contents_dir["starting_sector"] = self.contents_starting_sector
        self.contents_dir["stream_size"] = self.payload_len
        self.ole_native_dir["stream_size"] = self.payload_len
        self.ole_native_dir["starting_sector"] = self.contents_starting_sector

        #Fill in header fields
        self.num_difat_secs = 0
        if self.num_difat_secs == 0:
            self.header_fields["first_msat_secid"] = OLEFile.ENDOFCHAIN
        else:
            self.header_fields["first_msat_secid"] = self.header_chain_entries + self.initial_offset
        self.header_fields["first_ssat_secid"] = 2
        self.header_fields["num_dir_sectors"] = 0 #Num dir sectors  - must be zeros

        #Build SSAT
        self.ssat_table = [1, OLEFile.ENDOFCHAIN]
        for i in range(3, self.num_short_sectors):
            self.ssat_table.append(i)
        self.ssat_table.append(OLEFile.ENDOFCHAIN)
        while len(self.ssat_table) < 128:
            self.ssat_table.append(OLEFile.FREESECT)

        #Make the sat up until the next msat
        self.sat.append(OLEFile.FAT_SECT) #FAT sector
        self.sat.append(OLEFile.ENDOFCHAIN) #DIR entry
        self.sat.append(OLEFile.ENDOFCHAIN) #DIR entry
        start_index = 4
        for i in range(start_index, start_index - 1 + math.ceil(self.root_dir["stream_size"] / 512)):
            self.sat.append(i)
        self.sat.append(OLEFile.ENDOFCHAIN)
        # print(len(self.sat), sat_end)
        sat_end = max(self.sec_ids_per_msat + 1, 1 + int(math.ceil(start_index / self.sec_ids_per_msat)) * self.sec_ids_per_msat)
        while len(self.sat) < sat_end - 1:
            self.sat.append(OLEFile.FREESECT)

        #Make header MSAT
        self.header_msat.append(0) #This zero for the header
        for sec_id in range(self.start_sec_id, 0):
            self.header_msat.append(sec_id)
        while len(self.header_msat) <= self.header_chain_entries - 1:
            self.header_msat.append(OLEFile.FREESECT)
        self.all_msats.append(self.header_msat)

        #Make chain
        self.chain = []

    def build_sat_and_msat_long_file(self):
        #Make the sat up until the next msat
        self.sat.append(OLEFile.FAT_SECT) #FAT sector
        self.sat.append(OLEFile.ENDOFCHAIN) #DIR entry
        self.sat.append(OLEFile.ENDOFCHAIN) #DIR entry
        self.sat.append(OLEFile.ENDOFCHAIN) #DIR entry
        self.num_difat_secs = 0
        i = 0
        while i < (self.num_difat_secs + math.ceil((self.num_sectors + 2.5 + self.num_difat_secs) / 127) - 1):
            if (len(self.sat) - (self.header_chain_entries + self.initial_offset)) % self.sec_ids_per_msat == 0:
                self.sat.append(OLEFile.DIFAT_SECT)
                self.num_difat_secs += 1
            else:
                self.sat.append(OLEFile.FAT_SECT) #FAT sector
            i += 1
        if (len(self.sat) - (self.header_chain_entries + self.initial_offset)) % self.sec_ids_per_msat == 0:
            self.sat.append(OLEFile.DIFAT_SECT)
            self.num_difat_secs += 1
        print("num_difat_secs", self.num_difat_secs, i, len(self.sat) + 1)
        self.num_fat_sectors = math.ceil((self.num_sectors + 2.5 + self.num_difat_secs) / 127)
        print("total_secs_msat", self.total_secs_msat, self.num_sectors, self.num_difat_secs)

        #Make the constants and size dependent calcs
        total_fat_secs = math.ceil((self.num_sectors + 2.5 + self.num_difat_secs) / 127)
        chain_end = 1 + int(self.num_sectors + 2 + self.num_difat_secs + total_fat_secs) #Correct
        start_index = self.initial_offset + total_fat_secs + 1
        sat_end = max(self.sec_ids_per_msat + 1, 1 + int(math.ceil(start_index / self.sec_ids_per_msat)) * self.sec_ids_per_msat)
        print("chain end", chain_end, self.num_sectors, total_fat_secs, self.num_difat_secs, start_index, sat_end)

        #Make header MSAT
        self.header_msat.append(0) #This zero for the header
        msat_end = math.ceil(chain_end / 128) + 3
        for sec_id in range(self.start_sec_id, min(msat_end, self.start_sec_id + self.header_chain_entries - 1)):
            self.header_msat.append(sec_id)
        while len(self.header_msat) <= self.header_chain_entries - 1:
            self.header_msat.append(OLEFile.FREESECT)
        self.all_msats.append(self.header_msat)

        self.contents_starting_sector = msat_end
        self.ole_native_start = self.contents_starting_sector 
        self.last_msat_sec_id = self.start_sec_id + min(math.ceil(chain_end / 128), self.header_chain_entries - 1)

        # print(self.sat)
        #Fill with remaining sec ids
        start_index = len(self.sat) + 1
        for i in range(start_index, min(chain_end, sat_end)):
            self.sat.append(i)

        self.chain = []
        if chain_end < sat_end:
            #Fill with 0s
            self.sat.append(OLEFile.ENDOFCHAIN)
            # print(len(self.sat), sat_end)
            while len(self.sat) < sat_end - 1:
                self.sat.append(OLEFile.FREESECT)
        else:
            #Build the sec id chain
            print("chain going from {} to {}".format(sat_end, chain_end))
            start_index = len(self.sat) + 1
            final_msat = chain_end - self.num_sectors
            self.next_last_msat_sec_id = final_msat
            self.is_last_msat = False
            self.added_msat = False
            for i in range(start_index, chain_end):
                self.chain.append(i)
                if (i - 14080) % 16256 == 0:
                    curr_msat = []
                    curr_msat.append(self.last_msat_sec_id)
                    self.added_msat = True
                    self.is_last_msat = final_msat < (self.last_msat_sec_id + self.sec_ids_per_msat - 2)
                    if not self.is_last_msat:
                        self.next_last_msat_sec_id = self.last_msat_sec_id + self.sec_ids_per_msat
                    else:
                        self.next_last_msat_sec_id = final_msat
                    print(self.is_last_msat, self.last_msat_sec_id, self.next_last_msat_sec_id, total_fat_secs, (start_index - 1) - self.last_msat_sec_id)
                    for sec_id in range(self.last_msat_sec_id + 2, self.next_last_msat_sec_id):
                        curr_msat.append(sec_id)

                    #Fill in the last MSAT with freesect and then an end of chain
                    while len(curr_msat) < self.sec_ids_per_msat - 1:
                        curr_msat.append(OLEFile.FREESECT)

                    if self.is_last_msat:
                        curr_msat.append(OLEFile.ENDOFCHAIN)
                        self.ole_native_start = self.next_last_msat_sec_id
                        self.last_msat_sec_id = self.next_last_msat_sec_id
                    else:
                        curr_msat.append(self.next_last_msat_sec_id + 1)
                        self.last_msat_sec_id = self.next_last_msat_sec_id
                    self.chain.extend(curr_msat)
                    self.all_msats.append(curr_msat)
            #Append the end of chain and fill with 0
            self.chain.append(OLEFile.ENDOFCHAIN)
            chain_end_len = int(math.ceil((len(self.chain) + len(self.sat)) / self.sec_ids_per_msat) * self.sec_ids_per_msat)
            while len(self.chain) + len(self.sat) < chain_end_len:
                self.chain.append(OLEFile.FREESECT)

            if not self.is_last_msat and self.added_msat:
                curr_msat = []
                curr_msat.append(self.last_msat_sec_id)

                #Fill in the last MSAT with freesect and then an end of chain
                while len(curr_msat) < self.sec_ids_per_msat - 1:
                    curr_msat.append(OLEFile.FREESECT)

                curr_msat.append(OLEFile.ENDOFCHAIN)
                self.next_last_msat_sec_id = final_msat
                self.total_secs_msat += 1
                self.ole_native_start = self.next_last_msat_sec_id
                self.last_msat_sec_id = self.next_last_msat_sec_id
                self.chain.extend(curr_msat)
                self.all_msats.append(curr_msat)

        print("chain_end_len", len(self.chain), len(self.chain) % self.sec_ids_per_msat, len(self.sat), len(self.sat) % self.sec_ids_per_msat)
        #Fill in dir fields
        self.contents_dir["starting_sector"] = self.contents_starting_sector
        self.contents_dir["stream_size"] = self.payload_len
        self.ole_native_dir["stream_size"] = self.payload_len
        self.ole_native_dir["starting_sector"] = self.ole_native_start

        #Fill in header fields
        if self.num_difat_secs == 0:
            self.header_fields["first_msat_secid"] = OLEFile.ENDOFCHAIN
        else:
            self.header_fields["first_msat_secid"] = self.header_chain_entries + self.initial_offset
        self.header_fields["first_ssat_secid"] = 2
        self.header_fields["num_dir_sectors"] = 0 #Num dir sectors  - must be zeros
        print("ole_native_start", self.ole_native_start)

    def write_header(self, outfile):
        self.write_bytes(outfile, self.header_fields["magic"])
        self.write_bytes(outfile, self.header_fields["uuid"])
        self.write_bytes(outfile, self.header_fields["rev_num"])
        self.write_bytes(outfile, self.header_fields["version_num"])
        self.write_bytes(outfile, self.header_fields["endianness"])
        self.write_bytes(outfile, self.header_fields["sec_size"])
        self.write_bytes(outfile, self.header_fields["ssec_size"])
        self.write_bytes(outfile, self.header_fields["reserved"])
        # print(self.cur_filelen, "num_dir_sectors", self.header_fields["num_dir_sectors"])
        self.write_int(outfile, self.header_fields["num_dir_sectors"], num_bytes=4)
        print(self.cur_filelen, "num_fat_sectors", self.num_fat_sectors)
        self.write_int(outfile, self.num_fat_sectors, num_bytes=4)
        print(self.cur_filelen, "first_sec_id", self.header_fields["first_sec_id"])
        self.write_int(outfile, self.header_fields["first_sec_id"], num_bytes=4)
        self.write_bytes(outfile, self.header_fields["not_used"])
        self.write_bytes(outfile, self.header_fields["cutoff_size"])
        print(self.cur_filelen, "first_ssat_secid", self.header_fields["first_ssat_secid"])
        self.write_int(outfile, self.header_fields["first_ssat_secid"], num_bytes=4)
        print(self.cur_filelen, "total_secs_ssat", self.header_fields["total_secs_ssat"])
        self.write_int(outfile, self.header_fields["total_secs_ssat"], num_bytes=4)
        print(self.cur_filelen, "first_msat_secid", self.header_fields["first_msat_secid"])
        self.write_int(outfile, self.header_fields["first_msat_secid"], num_bytes=4)
        #This points to the next MSAT
        print(self.cur_filelen, "total_secs_msat", self.num_difat_secs)
        self.write_int(outfile, self.num_difat_secs, num_bytes=4)
        # print(self.cur_filelen, "msat_start", self.header_msat)
        for val in self.header_msat:
            self.write_int(outfile, val, num_bytes=4)
        
    def make_payload_prefix(self, actual_payload_len):
        print("actual_payload_len", actual_payload_len)
        self.payload_prefix_len = 0

        # self.payload_header_start_len = actual_payload_len - 4
        self.payload_prefix_len += 4

        self.payload_header_val_1 = 2
        self.payload_prefix_len += 2

        # self.display_filename = display_filename
        print("display_filename", self.display_filename)
        self.payload_prefix_len += (len(self.display_filename) + 1)

        # self.long_filename = "C:\\Users\\12149\\{}".format(self.display_filename)
        self.payload_prefix_len += (len(self.long_filename) + 1)

        self.payload_header_val_2 = "00 00 03 00"
        self.payload_prefix_len += 4

        self.payload_header_temp_len = (len(self.temp_filename) + 1)
        self.payload_prefix_len += 4

        # self.temp_filename = "C:\\Users\\12149\\AppData\\Local\\Temp\\{1AC28F20-7FA0-4EC6-B88E-C4E6BBD2A693}\\{}".format(self.display_filename)
        self.payload_prefix_len += (len(self.temp_filename) + 1)

        self.payload_prefix_len += 4
        self.payload_prefix_some_unknown_len = actual_payload_len
        self.payload_header_start_len = actual_payload_len + (3*(self.payload_prefix_len - 17) + 17)

        self.payload_suffix_len = 4 + 2 * len(self.temp_filename) + 4 + 2 * len(self.display_filename) + 4 + 2 * len(self.long_filename)

    def read_payload(self):
        with open(self.bin_filename, "rb") as infile:
            self.payload = infile.read()
            print("add_payload_prefix", self.add_payload_prefix)
            if self.add_payload_prefix:
                self.make_payload_prefix(len(self.payload))
                self.payload_len = len(self.payload) + self.payload_prefix_len + self.payload_suffix_len
            else:
                self.payload_len = len(self.payload)

    def write_payload(self, outfile):
        if self.add_payload_prefix:
            self.write_int(outfile, self.payload_header_start_len, num_bytes=4)
            self.write_int(outfile, self.payload_header_val_1, num_bytes=2)
            self.write_str(outfile, self.display_filename)
            self.write_str(outfile, self.long_filename)
            self.write_bytes(outfile, self.payload_header_val_2)
            self.write_int(outfile, self.payload_header_temp_len, num_bytes=4)
            self.write_str(outfile, self.temp_filename)
            self.write_int(outfile, self.payload_prefix_some_unknown_len, num_bytes=4)

        outfile.write(self.payload)
        self.cur_filelen += len(self.payload)

        if self.add_payload_prefix:
            self.write_int(outfile, self.payload_header_temp_len - 1, num_bytes=4)
            for char in self.temp_filename:
                self.write_str(outfile, char)
            self.write_int(outfile, len(self.display_filename), num_bytes=4)
            for char in self.display_filename:
                self.write_str(outfile, char)
            self.write_int(outfile, len(self.long_filename), num_bytes=4)
            for char in self.long_filename:
                self.write_str(outfile, char)


    def write_str(self, outfile, ascii_str):
        outfile.write(bytes(ascii_str + '\0', "ascii"))
        self.cur_filelen += (len(ascii_str) + 1)
    
    def write_bytes(self, outfile, bytes_str):
        out_bytes = bytes_str.upper().split(" ")
        for byte in out_bytes:
            outfile.write(binascii.unhexlify(byte))
        self.cur_filelen += len(out_bytes)

    def write_int(self, outfile, val, num_bytes=4):
        outfile.write(val.to_bytes(num_bytes, byteorder=self.byteorder, signed=True))
        self.cur_filelen += num_bytes

    def write_dir_entry(self, outfile, dir_entry):
        #Write name 2 bytes per char        64 bytes
        name_len = len(dir_entry["name"].split(" "))
        name = dir_entry["name"] + (64 - name_len) * " 00"
        self.write_bytes(outfile, name)
        #Write len of name + 1              2 bytes
        self.write_bytes(outfile, dir_entry["name_len"])
        #Obj type - root 05, unknown/unallocated 00, storage 01, stream 02  1 byte
        self.write_bytes(outfile, dir_entry["type"])
        #Color- 00 red, 01 black                        1 byte
        self.write_bytes(outfile, dir_entry["color"])
        #Left sibling - 0xFFFFFFFF for no sibling       4 bytes
        self.write_bytes(outfile, dir_entry["left"])
        #Right sibling - 0xFFFFFFFF for no sibling      4 bytes
        self.write_bytes(outfile, dir_entry["right"])
        #Child id - 0xFFFFFFFF for no child             4 bytes
        self.write_bytes(outfile, dir_entry["child"])
        #Class id - zeros                               16 bytes
        if "uuid" in dir_entry:
            self.write_bytes(outfile, dir_entry["uuid"])
        else:
            self.write_bytes(outfile, (" 00" * 16).strip())
        #State bits - zeros                             4 bytes
        self.write_bytes(outfile, (" 00" * 4).strip())
        #CTIME - zeros                                  8 bytes
        self.write_bytes(outfile, (" 00" * 8).strip())
        #MTIME - zeros                                  8 bytes
        self.write_bytes(outfile, (" 00" * 8).strip())
        #Starting sector - for root first sector of ministream, storage is all zeros        4 bytes
        if type(dir_entry["starting_sector"]) is str:
            self.write_bytes(outfile, dir_entry["starting_sector"])
        else:
            self.write_int(outfile, dir_entry["starting_sector"], num_bytes=4)
        #Stream size - stream = size of user data, root=size of mini stream, storage=all zeros  8 bytes
        if type(dir_entry["stream_size"]) is str:
            self.write_bytes(outfile, dir_entry["stream_size"])
        else:
            self.write_int(outfile, dir_entry["stream_size"], num_bytes=8)

    def write(self, out_filename):
        with open(out_filename, "wb") as outfile:
            #Build the file structure
            self.read_payload()
            self.build_sat_and_msat()
            
            self.cur_filelen = 0
            #Write the header
            self.write_header(outfile)
            
            #Write the SAT table
            if len(self.sat) > self.sec_ids_per_msat:
                for val in self.sat[:self.sec_ids_per_msat]:
                    self.write_int(outfile, val, num_bytes=4)
            else:
                for val in self.sat:
                    self.write_int(outfile, val, num_bytes=4)

            #Write the dir entries
            for dir_entry in self.dir_entries:
                self.write_dir_entry(outfile, dir_entry)

            #Write the SSAT table
            # self.write_bytes(outfile, self.ssat_table)
            for val in self.ssat_table:
                self.write_int(outfile, val, num_bytes=4)

            #Write the prefix
            if self.payload_len >= 1024:
                self.write_bytes(outfile, self.prefix)
            else:
                self.write_bytes(outfile, " ".join(self.prefix.split(" ")[:128]))

            #Write the SAT table remainder
            # print("chain", self.sat[self.sec_ids_per_msat:])
            if len(self.sat) > self.sec_ids_per_msat:
                for val in self.sat[self.sec_ids_per_msat:]:
                    self.write_int(outfile, val, num_bytes=4)

            #Write the chain
            # print("chain", self.chain)
            for val in self.chain:
                self.write_int(outfile, val, num_bytes=4)

            #Write the payload
            self.write_payload(outfile)

            #Write the remainder
            end_len = int(math.ceil(self.cur_filelen / self.sector_size)) * self.sector_size
            while end_len != self.cur_filelen:
                self.write_int(outfile, 0, num_bytes=1)

class ZipOLEFile(OLEFile):
    #This might be the general class
    def __init__(self, bin_filename, display_filename, add_payload_prefix=True):
        super().__init__(bin_filename, display_filename, add_payload_prefix=add_payload_prefix)
        self.root_dir["child"] = "01 00 00 00"
        self.root_dir["uuid"] = "0C 00 03 00 00 00 00 00 C0 00 00 00 00 00 00 46"
        self.root_dir["stream_size"] = 128

        self.comp_obj_dir["left"] = "FF FF FF FF"
        self.comp_obj_dir["right"] = "02 00 00 00"
        self.comp_obj_dir["child"] = "FF FF FF FF"
        self.comp_obj_dir["starting_sector"] = 0
        self.comp_obj_dir["stream_size"] = 76
        #                    400            480                500                  580
        self.dir_entries = [self.root_dir, self.comp_obj_dir, self.ole_native_dir, self.empty_dir]
        self.ssat_table = [1, OLEFile.ENDOFCHAIN] + [OLEFile.FREESECT]*126
        self.prefix = "01 00 fe ff 03 0a 00 00 ff ff ff ff 0c 00 03 00" + \
            " 00 00 00 00 c0 00 00 00 00 00 00 46 0c 00 00 00" + \
            " 4f 4c 45 20 50 61 63 6b 61 67 65 00 00 00 00 00" + \
            " 08 00 00 00 50 61 63 6b 61 67 65 00 f4 39 b2 71" + \
            " 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" * 28

class PDFOLEFile(OLEFile):
    def __init__(self, bin_filename, display_filename, version, add_payload_prefix=True):
        super().__init__(bin_filename, display_filename, add_payload_prefix=add_payload_prefix)
        self.root_dir["child"] = "02 00 00 00"
        self.root_dir["uuid"] = "65 CA 01 B8 FC A1 D0 11 85 AD 44 45 53 54 00 00"
        self.root_dir["stream_size"] = 192
        
        self.comp_obj_dir["left"] = "01 00 00 00"
        self.comp_obj_dir["right"] = "03 00 00 00"
        self.comp_obj_dir["child"] = "FF FF FF FF"
        self.comp_obj_dir["starting_sector"] = 1
        #                   400             480          500                580
        self.dir_entries = [self.root_dir, self.ole_dir, self.comp_obj_dir, self.contents_dir]
        # self.ssat_table = "FE FF FF FF 02 00 00 00 FE FF FF FF FF FF FF FF" + " FF" * 496
        self.ssat_table = [OLEFile.ENDOFCHAIN, 2, OLEFile.ENDOFCHAIN] + [OLEFile.FREESECT]*125

        self.prefix = "01 00 00 02 00 00 00 00 00 00 00 00 00 00 00 00" + \
            " 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + \
            " 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + \
            " 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" + \
            " 01 00 FE FF 03 0A 00 00 FF FF FF FF 65 CA 01 B8" + \
            " FC A1 D0 11 85 AD 44 45 53 54 00 00 11 00 00 00" + \
            " 41 63 72 6F 62 61 74 20 44 6f 63 75 6D 65 6E 74"
        if version == 1:
            self.prefix += " 00 00 00 00 00 14 00 00 00 41 63 72 6F 62 61 74" + \
            " 2E 44 6F 63 75 6D 65 6E 74 2E 44 43 00 F4 39 B2" + \
            " 71 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"
            self.comp_obj_dir["stream_size"] = 93
        else:
            self.prefix += " 00 00 00 00 00 15 00 00 00 41 63 72 6F 45 78 63" + \
            " 68 2E 44 6F 63 75 6D 65 6E 74 2E 44 43 00 F4 39" + \
            " B2 71 00 00 00 00 00 00 00 00 00 00 00 00 00 00"
            self.comp_obj_dir["stream_size"] = 94

        self.prefix += " 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" * 22
