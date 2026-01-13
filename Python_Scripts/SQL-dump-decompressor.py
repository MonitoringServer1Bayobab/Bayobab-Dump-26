import os
import gzip
import shutil

def decompress_gz_to_sql(file_name):
    input_path = os.path.join(os.path.dirname(__file__), file_name)
    output_path = os.path.splitext(input_path)[0]  # Remove the '.gz' extension to get .sql
 
    # Decompress the file
    with gzip.open(input_path, 'rb') as f_in:
        with open(output_path, 'wb') as f_out:
            shutil.copyfileobj(f_in, f_out)
   
 
decompress_gz_to_sql(r'C:\Users\lynette.mutuku\Documents\ett3\dump.sql.gz')