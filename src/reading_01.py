import os
import json
from datetime import datetime
from config_tools import DIR


class FileMetadataCache:
    """Manages the file metadata cache for efficient XML processing."""

    def __init__(self, cache_file):
        self.cache_file = cache_file
        self.cache_data = self.load_cache()

    def load_cache(self):
        """Load the cache data from a JSON file."""
        if os.path.exists(self.cache_file):
            with open(self.cache_file, 'r') as f:
                return json.load(f)
        return {}

    def save_cache(self):
        """Save the current cache data to a JSON file."""
        with open(self.cache_file, 'w') as f:
            json.dump(self.cache_data, f, indent=4)

    def is_file_new_or_modified(self, file_name, timestamp):
        """Check if the file is new or has been modified."""
        return file_name not in self.cache_data or self.cache_data[file_name] != timestamp

    def update_cache(self, file_name, timestamp):
        """Update the cache with the latest timestamp for the file."""
        self.cache_data[file_name] = timestamp


class XMLFileProcessor:
    """Processes XML files and handles caching of file metadata."""

    def __init__(self, directory, cache_file):
        self.directory = directory
        self.cache = FileMetadataCache(cache_file)

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
                    file_path = os.path.join(root, file)
                    file_name, timestamp = self.get_file_metadata(file_path)
                    if file_name and self.cache.is_file_new_or_modified(file_name, timestamp):
                        new_files.append({'file_name': file_name, 'file_path': file_path, 'timestamp': timestamp})
                        self.cache.update_cache(file_name, timestamp)
        return new_files

    def process_new_files(self, json_file):
        """Process new or modified XML files and save the details to a JSON file."""
        new_files = self.scan_for_new_files()
        if new_files:
            with open(json_file, 'w') as f:
                json.dump(new_files, f, indent=4)
            print(f"New file details saved to: {json_file}")
        else:
            print("No new or modified files found.")
        self.cache.save_cache()


# Example usage:

path_to = DIR()
processor = XMLFileProcessor(path_to.xml_files,path_to.cache_file)
processor.process_new_files(path_to.json_object)
