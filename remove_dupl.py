import os
import hashlib
import collections


def compute_partial_hash_of(file_path, seek_pos, chunk_size):
    with open(file_path, "rb") as f:
        f.seek(seek_pos)
        data = f.read(chunk_size)
        if not data:
            return  # no more yield
        hashes = [hashlib.sha256(), hashlib.md5()]
        hashes[0].update(data)
        hashes[1].update(data)
        yield "_".join(map(lambda y: y.hexdigest(), hashes))


def make_challenge(file_path):
    chunk_size = 1048576
    file_size = os.path.getsize(file_path)
    yield file_size
    yield from compute_partial_hash_of(file_path, 0, chunk_size)
    yield from compute_partial_hash_of(file_path, file_size - chunk_size, chunk_size)


def fuzzy_group_files(target_directory):
    target_files = []
    for root, _, files in os.walk(target_directory):
        target_files.extend(map(lambda x: os.path.join(root, x), files))

    groups, challenge, challenge_result = init_challenge(target_files)

    found_challenge = True
    while found_challenge:
        found_challenge = False
        for _, v in groups.items():
            if len(v) > 1:
                for file_path in v:
                    if challenge[file_path] is not None:
                        try:
                            challenge_result[file_path].append(
                                next(challenge[file_path]))
                            found_challenge = True
                        except StopIteration:
                            challenge[file_path] = None
                            pass
        groups = regroup(challenge_result)
        print(collections.Counter(map(lambda x: len(x), groups.values())))
    return groups


def regroup(challenge_result):
    groups = {}  # Regroup items
    for file_path, v in challenge_result.items():
        key = "/".join(map(lambda x: str(x), v))
        if key not in groups:
            groups[key] = []
        groups[key].append(file_path)
    return groups


def init_challenge(target_files):
    groups = {"": []}
    challenge = {}
    challenge_result = {}
    for file_path in target_files:
        challenge[file_path] = make_challenge(file_path)
        challenge_result[file_path] = []
        groups[""].append(file_path)
    return groups, challenge, challenge_result


def main():
    target_directory = r"C:\Users\SL\Desktop\test"
    groups = fuzzy_group_files(target_directory)
    for _, v in groups.items():
        if len(v) > 1:
            print("### Group ###")
            print("\n".join(v))


if __name__ == "__main__":
    main()
