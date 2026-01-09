from transmission_rpc import Client

# c = Client(protocol='https', host="tr4-1.404.svsoft.fun", port=40000, username="admin", password="Wp8004Wp8004",verify=False)
c = Client(host="10.0.4.20", port=9091, username="admin", password="Wp@8004Wp@8004")
# c.change_torrent(ids=963, tracker_replace=[(982, "https://tracker.pterclub.net/announce?passkey=d47e62cfabe5f92331623589f060b00d")])
# c.change_torrent(ids=963, tracker_add=["https://tracker.pterclub.net/announce?passkey=d47e62cfabe5f92331623589f060b00d"])
# c.change_torrent(ids=963, tracker_remove=[982])
tor = [2245, 2248, 2249, 2271, 2272, 2275, 2278, 2279, 2303, 2319, 2320, 2321, 2322, 2323, 2328, 2330, 2331, 2342, 2344, 2345, 2356, 2357, 2359, 2360, 2361, 2362, 2373, 2374, 2375,
       2376, 2390, 2411, 5322, 5323, 5700, 6591, 6592, 6593, 6594, 6595, 6596, 6597, 6598, 6599, 6600, 6601, 6602, 6603, 6604, 6605, 6606, 6607, 6608, 6609, 6610, 6611, 6612, 6613,
       6614, 6615, 9731]
torrents = c.get_torrents()


def change_tracker(tors):
    for t in tors:
        trackers = t.get('trackers')
        for tracker in trackers:
            c.change_torrent(ids=t.get('id'), tracker_remove=[tracker.get('id')])
        c.change_torrent(ids=t.get('id'), tracker_add=["https://tracker.rousipt.com/announce.php?passkey=d2ec55af42b9bf1b29c1903646407be3"])


def change_path(tors):
    target_dir = '/downloads/disk4/PT'
    for t in tors:
        if t.download_dir == target_dir:
            print(t.get('name'))


change_path(torrents)
pass
