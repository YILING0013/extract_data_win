import sys
import gzip
import json
from PIL import Image

def bits_to_int(bits):
    return int("".join(map(str, bits)), 2)

def extract_bits_from_pixel(pixel_list):
    bits_list = []
    for pixel in pixel_list:
        if pixel[3] & 1:
            bits_list.append(1)
        else:
            bits_list.append(0)
    return bits_list

def read_hidden_data_from_image(input_path):
    img = Image.open(input_path)
    img = img.transpose(5)
    pixel_list = list(img.getdata())

    bits_list = extract_bits_from_pixel(pixel_list)

    prefix_length = len("stealth_pngcomp") * 8
    length_bits = bits_list[prefix_length : prefix_length + 32]
    data_length = bits_to_int(length_bits)

    data_bits = bits_list[prefix_length + 32 : prefix_length + 32 + data_length]
    data_bytes = bytearray()
    for i in range(0, len(data_bits), 8):
        byte = data_bits[i : i + 8]
        data_bytes.append(int("".join(map(str, byte)), 2))

    decompressed_data = gzip.decompress(data_bytes).decode("utf-8")
    d = json.loads(decompressed_data)
    comment_data = json.loads(d.get("Comment", "{}"))
    return comment_data

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: extract_data.exe <input_image_path> <output_json_path>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        comment_data = read_hidden_data_from_image(input_path)
        with open(output_path, 'w') as outfile:
            json.dump(comment_data, outfile)
        print(f"Data successfully extracted to {output_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)
