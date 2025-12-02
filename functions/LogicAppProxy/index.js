const axios = require("axios");

module.exports = async function (context, req) {
    context.log("LogicAppProxy trigger:", req.method, req.url);

    const corsHeaders = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Origin, X-Requested-With, Content-Type, Accept, Authorization"
    };

    // Handle OPTIONS preflight
    if (req.method === "OPTIONS") {
        context.res = {
            status: 200,
            headers: {
                ...corsHeaders,
                "Content-Type": "application/json"
            }
        };
        return;
    }

    try {
        const body = req.body || {};
        const { type } = body;

        // Missing type
        if (!type) {
            context.res = {
                status: 400,
                headers: { ...corsHeaders, "Content-Type": "application/json" },
                body: { ok: false, error: "Missing type" }
            };
            return;
        }

        // ---------------------------------------------------------------
        // ⭐ UID LOOKUP (RESTORED EXACT BEHAVIOR + CONTENT-TYPE FIX)
        // ---------------------------------------------------------------
        if (String(type).trim().toUpperCase() === "UID") {
            const { uid } = body;

            if (!uid) {
                context.res = {
                    status: 400,
                    headers: { ...corsHeaders, "Content-Type": "application/json" },
                    body: { ok: false, error: "Missing uid" }
                };
                return;
            }

            const url = process.env.LOGICAPP_UID_URL + `&UID=${encodeURIComponent(uid)}`;
            context.log("Calling UID Logic App:", url);

            const logicResponse = await axios.get(url, { timeout: 0 });

            context.res = {
                status: 200,
                headers: {
                    ...corsHeaders,
                    "Content-Type": "application/json"  // REQUIRED so React parses JSON correctly
                },
                body: logicResponse.data  // return raw JSON exactly as before
            };
            return;
        }

        // ---------------------------------------------------------------
        // ⭐ VSO REQUESTS (Stage-based, POST-only, Fiber Util integration)
        // ---------------------------------------------------------------
        if (String(type).trim().toUpperCase() === "VSO") {
            const url = process.env.LOGICAPP_VSO_URL;

            if (!url) {
                context.res = {
                    status: 500,
                    headers: { ...corsHeaders, "Content-Type": "application/json" },
                    body: { ok: false, error: "LogicApp VSO URL missing" }
                };
                return;
            }

            // Ensure Stage is ALWAYS string (Logic App schema requires string)
            if (body.Stage !== undefined) {
                body.Stage = String(body.Stage);
            }

            // Ensure Tags is sent as a string (Logic App schema expects a string)
            if (body.Tags !== undefined) {
                try {
                    if (Array.isArray(body.Tags)) {
                        // Join arrays into a semicolon-separated string (matches UI formatting)
                        body.Tags = body.Tags.join('; ');
                    } else if (typeof body.Tags === 'object' && body.Tags !== null) {
                        // If it's an object, stringify it to preserve content
                        body.Tags = JSON.stringify(body.Tags);
                    } else {
                        body.Tags = String(body.Tags || '');
                    }
                } catch (e) {
                    // Fallback to a safe string conversion
                    body.Tags = String(body.Tags);
                }
            }

            context.log("Calling VSO Logic App (POST):", url);

            const logicResponse = await axios.post(url, body, {
                timeout: 500000
            });

            context.res = {
                status: 200,
                headers: {
                    ...corsHeaders,
                    "Content-Type": "application/json"
                },
                body: logicResponse.data
            };
            return;
        }

        // Unknown type
        context.res = {
            status: 400,
            headers: { ...corsHeaders, "Content-Type": "application/json" },
            body: { ok: false, error: `Unsupported type: ${type}` }
        };

    } catch (err) {
        context.log.error("Proxy error:", err);

        context.res = {
            status: 500,
            headers: { ...corsHeaders, "Content-Type": "application/json" },
            body: {
                ok: false,
                error: err.message || "Unknown error",
                details: err.response?.data || null
            }
        };
    }
};
