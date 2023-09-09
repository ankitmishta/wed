const mongoose = require('mongoose');
var validator = require("validator")

const contactSchema = new mongoose.Schema(
    {
        name: {
            type: String,
            required :true,
            minlength: [3, "Enter atleast 3 character"]
        },
        email: {
            type: String,
            required :true,
            validate(value){
                if(!validator.isEmail(value)){
                    throw new Error("Email is Invalid")
                }
            }
        },
        phone: {
            type: String,
            required :true,
            validate(phone){
                if(!validator.isMobilePhone(phone)){
                    throw new Error("provide correct phone number")
                }
            }
        },
        proof: {
            type: String,
            required :true,
        },
        business: {
            type: String,
            required :true
        },
        sales: {
            type: String,
            required :true
        },
        gstn: {
            type: String,
            required :true
        }

    }
);

const contactInfo = mongoose.model('contactInfo',contactSchema);
module.exports = contactInfo;