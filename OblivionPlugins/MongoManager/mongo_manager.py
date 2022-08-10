import pymongo


class MongoManager:
    def __init__(self, report_folder, mongo_address=None):
        self.report_folder = report_folder
        self.mongo_address = mongo_address if mongo_address is not None else "mongodb://localhost:27017"
