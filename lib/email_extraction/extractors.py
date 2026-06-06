import os
import glob
from typing import List, Optional, Callable, Dict

import extract_msg


def get_msg_files(directory: str, recursive: bool = False) -> List[str]:
    """
    Get list of .msg files in a directory.
    
    Args:
        directory: Path to directory containing .msg files
        recursive: Whether to search subdirectories recursively
        
    Returns:
        List of full paths to .msg files
    """
    pattern = "**/*.msg" if recursive else "*.msg"
    return glob.glob(os.path.join(directory, pattern), recursive=recursive)


def extract_attachments_from_msg(
    msg_file_path: str,
    dest_folder: str,
    filename_filter: Optional[Callable[[str], bool]] = None,
    verbose: bool = True
) -> List[str]:
    """
    Extract attachments from a single .msg file.
    
    Args:
        msg_file_path: Path to the .msg file
        dest_folder: Destination folder for extracted attachments
        filename_filter: Optional callable that takes filename and returns True if should be extracted
        verbose: Whether to print extraction status messages
        
    Returns:
        List of paths to extracted files
    """
    os.makedirs(dest_folder, exist_ok=True)
    extracted_files = []
    
    msg = extract_msg.Message(msg_file_path)
    
    for attachment in msg.attachments:
        filename = attachment.longFilename or attachment.shortFilename
        
        if not filename:
            continue
            
        # Apply filter if provided
        if filename_filter and not filename_filter(filename):
            continue
        
        output_path = os.path.join(dest_folder, filename)
        if os.path.exists(output_path):
            base, ext = os.path.splitext(filename)
            counter = 1
            while True:
                candidate = os.path.join(dest_folder, f"{base}_{counter}{ext}")
                if not os.path.exists(candidate):
                    output_path = candidate
                    break
                counter += 1

        with open(output_path, "wb") as f:
            f.write(attachment.data)
        
        extracted_files.append(output_path)
        
        if verbose:
            print(f"Saved: {output_path}")
    
    return extracted_files


def extract_attachments_from_multiple_msg(
    msg_files: List[str],
    dest_folder: str,
    filename_filter: Optional[Callable[[str], bool]] = None,
    verbose: bool = True
) -> Dict[str, List[str]]:
    """
    Extract attachments from multiple .msg files.
    
    Args:
        msg_files: List of paths to .msg files
        dest_folder: Destination folder for extracted attachments
        filename_filter: Optional callable that takes filename and returns True if should be extracted
        verbose: Whether to print extraction status messages
        
    Returns:
        Dictionary mapping msg file path to list of extracted attachment paths
    """
    results = {}
    
    for msg_file in msg_files:
        extracted = extract_attachments_from_msg(
            msg_file, 
            dest_folder, 
            filename_filter, 
            verbose
        )
        results[msg_file] = extracted
    
    return results



