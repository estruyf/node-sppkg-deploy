"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var crypto = require("crypto");
var uuidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
var uuid4 = (function () {
    function uuid4() {
    }
    uuid4.generate = function () {
        var rnd = crypto.randomBytes(16);
        rnd[6] = (rnd[6] & 0x0f) | 0x40;
        rnd[8] = (rnd[8] & 0x3f) | 0x80;
        var guid = rnd.toString('hex').match(/(.{8})(.{4})(.{4})(.{4})(.{12})/);
        guid.shift();
        return guid.join('-');
    };
    uuid4.valid = function (uuid) {
        return uuidPattern.test(uuid);
    };
    return uuid4;
}());
exports.default = uuid4;
//# sourceMappingURL=uuid4.js.map