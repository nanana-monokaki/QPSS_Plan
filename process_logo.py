from PIL import Image

def process_logo():
    in_path = "QPSSTOP.png"
    out_path = "QPSSTOP_transparent.png"
    
    img = Image.open(in_path).convert("RGBA")
    datas = img.getdata()
    
    new_data = []
    # Identify background color. usually white (255,255,255)
    # If the pixel is close to white, make it transparent
    for item in datas:
        # Check if it's already transparent
        if item[3] == 0:
            new_data.append(item)
            continue
            
        # Check if it's white or very close to white
        if item[0] > 240 and item[1] > 240 and item[2] > 240:
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
            
    img.putdata(new_data)
    img.save(out_path, "PNG")
    print(f"Saved transparent logo to {out_path}")

if __name__ == "__main__":
    process_logo()
