import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from PIL import Image
import io

def add_text_to_image(image, text_info_list):
    plt.figure(figsize=(10, 10))
    
    plt.imshow(image)
    ax = plt.gca()

    for text, position, font_size, font_name, color, style_properties in text_info_list:
        font_properties = fm.FontProperties(family=font_name, size=font_size,
                                            weight=style_properties['weight'],
                                            style=style_properties['style'])

        ax.text(position[0], position[1], text, fontproperties=font_properties,
                color=color, ha='left', va='top')

    plt.axis('off')
    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png', bbox_inches='tight', pad_inches=0)
    plt.close()

    img_buffer.seek(0)

    return img_buffer