import folium


def create_map_with_markers(coordinates, popups, colors):
    m = folium.Map(location=coordinates[0], zoom_start=10)

    for coord, text, color1 in zip(coordinates, popups, colors):
        folium.Marker(
            location=coord,
            popup=text,
            icon=folium.Icon(color=color1),
        ).add_to(m)

    m.save("map1.html")
