class UnsupportedMimetypeError(Exception):
    def __init__(self, observed_mimetype: str, file_path: str):
        super().__init__(f'Unsupported mimetype {observed_mimetype} for {file_path}')
        self.observed_mimetype = observed_mimetype
        self.file_path = file_path




