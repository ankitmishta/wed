const mongoose = require("mongoose");


const connectToMongo = async () =>{
    try {
        const conn = await mongoose.connect(process.env.mongoURI);
        console.log(`Connected To Database Successfully`);
    } catch (error) {
        console.log(`Connection Failed ${error}`);
    }
}

module.exports = connectToMongo;