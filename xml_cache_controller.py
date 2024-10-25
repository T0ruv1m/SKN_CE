import os
import csv
from datetime import datetime
from config_tools import DIR
import pandas as pd
import re

class CacheCreator:
    """Manages the file metadata cache for efficient XML processing."""

    def __init__(self, cache_file):
        self.cache_file = cache_file
        self.cache_data = self.load_cache()

    def load_cache(self):
        """Load the cache data from a CSV file."""
        cache_data = {}
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        cache_data[row['file_name']] = row['timestamp']
            except Exception as e:
                print(f"Error loading cache from CSV: {e}")
        return cache_data

    def save_cache(self):
        """Save the current cache data to a CSV file."""
        try:
            with open(self.cache_file, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['file_name', 'timestamp'])
                writer.writeheader()
                for file_name, timestamp in self.cache_data.items():
                    writer.writerow({'file_name': file_name, 'timestamp': timestamp})
        except Exception as e:
            print(f"Error saving cache to CSV: {e}")

    def is_file_new_or_modified(self, file_name, timestamp):
        """Check if the file is new or has been modified."""
        return file_name not in self.cache_data or self.cache_data[file_name] != timestamp

    def update_cache(self, file_name, timestamp):
        """Update the cache with the latest timestamp for the file."""
        self.cache_data[file_name] = timestamp

class XMLreading:
    """Processes XML files and handles caching of file metadata."""

    def __init__(self, directory, cache_file):
        self.directory = directory
        self.cache = CacheCreator(cache_file)

    def get_file_metadata(self, file_path):
        """Retrieve the last modification timestamp and file name of a file."""
        try:
            mtime = os.path.getmtime(file_path)
            timestamp = datetime.fromtimestamp(mtime).isoformat()
            file_name = os.path.basename(file_path)
            return file_name, timestamp
        except Exception as e:
            print(f"Error retrieving metadata for file {file_path}: {e}")
            return None, None

    def scan_for_new_files(self):
        """Scan the directory for new or modified XML files."""
        new_files = []
        for root, dirs, files in os.walk(self.directory):
            for file in files:
                if file.lower().endswith('.xml'):
                    # Skip files containing "-procEvento" followed by "NFe.xml" at the end
                    if re.search(r'-procEvento.*NFe\.xml$', file, re.IGNORECASE):
                        continue
                    
                    file_path = os.path.join(root, file)
                    file_name, timestamp = self.get_file_metadata(file_path)
                    if file_name and self.cache.is_file_new_or_modified(file_name, timestamp):
                        new_files.append({'file_name': file_name, 'file_path': file_path, 'timestamp': timestamp})
                        self.cache.update_cache(file_name, timestamp)
        return new_files

    def process_new_files(self, csv_file):
        """Process new or modified XML files and save the details to a CSV file."""
        new_files = self.scan_for_new_files()
        
        # Filter out files containing "-procEvento" in their names
        filtered_files = [
            file for file in new_files 
            if not re.search(r'-procEvento.*NFe\.xml$', file['file_name'], re.IGNORECASE)
        ]
        
        if filtered_files:
            try:
                with open(csv_file, 'w', newline='') as f:
                    writer = csv.DictWriter(f, fieldnames=['file_name', 'file_path', 'timestamp'])
                    writer.writeheader()
                    writer.writerows(filtered_files)
                print(f"New file details saved to: {csv_file}")
            except Exception as e:
                print(f"Error saving new files to CSV: {e}")
        else:
            print("No new or modified files found (excluding files with '-procEvento' in the name).")
        
        self.cache.save_cache()
    
class CacheOperations:
    """Manages cache-related tasks."""

    @staticmethod
    def clear_cache_files(cache_file):
        """Clear all cache files in the root folder, including merged_files.json."""
        try:
            if os.path.exists(cache_file):
                os.remove(cache_file)
                print(f"Deleted cache file: {cache_file}")
            else:
                print(f"Cache file not found: {cache_file}")

        except Exception as e:
            print(f"Error clearing cache files: {e}")


# Example usage:
if __name__ == "__main__":

    path_to = DIR()
    processor = XMLreading(path_to.xml_data, path_to.cache_compras)
    processor.process_new_files(path_to.new_compras)

    gestor_processor = XMLreading(path_to.gestor_data,path_to.cache_gestor)
    gestor_processor.process_new_files(path_to.new_gestor)