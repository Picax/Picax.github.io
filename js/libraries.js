/*! VelocityJS.org (1.5.0). (C) 2014 Julian Shapiro. MIT @license: en.wikipedia.org/wiki/MIT_License */
/*! VelocityJS.org jQuery Shim (1.0.1). (C) 2014 The jQuery Foundation. MIT @license: en.wikipedia.org/wiki/MIT_License. */
! function(a) {
    "use strict";

    function b(a) {
        var b = a.length,
            d = c.type(a);
        return "function" !== d && !c.isWindow(a) && (!(1 !== a.nodeType || !b) || ("array" === d || 0 === b || "number" == typeof b && b > 0 && b - 1 in a))
    }
    if (!a.jQuery) {
        var c = function(a, b) {
            return new c.fn.init(a, b)
        };
        c.isWindow = function(a) {
            return a && a === a.window
        }, c.type = function(a) {
            return a ? "object" == typeof a || "function" == typeof a ? e[g.call(a)] || "object" : typeof a : a + ""
        }, c.isArray = Array.isArray || function(a) {
            return "array" === c.type(a)
        }, c.isPlainObject = function(a) {
            var b;
            if (!a || "object" !== c.type(a) || a.nodeType || c.isWindow(a)) return !1;
            try {
                if (a.constructor && !f.call(a, "constructor") && !f.call(a.constructor.prototype, "isPrototypeOf")) return !1
            } catch (d) {
                return !1
            }
            for (b in a);
            return b === undefined || f.call(a, b)
        }, c.each = function(a, c, d) {
            var e = 0,
                f = a.length,
                g = b(a);
            if (d) {
                if (g)
                    for (; e < f && c.apply(a[e], d) !== !1; e++);
                else
                    for (e in a)
                        if (a.hasOwnProperty(e) && c.apply(a[e], d) === !1) break
            } else if (g)
                for (; e < f && c.call(a[e], e, a[e]) !== !1; e++);
            else
                for (e in a)
                    if (a.hasOwnProperty(e) && c.call(a[e], e, a[e]) === !1) break; return a
        }, c.data = function(a, b, e) {
            if (e === undefined) {
                var f = a[c.expando],
                    g = f && d[f];
                if (b === undefined) return g;
                if (g && b in g) return g[b]
            } else if (b !== undefined) {
                var h = a[c.expando] || (a[c.expando] = ++c.uuid);
                return d[h] = d[h] || {}, d[h][b] = e, e
            }
        }, c.removeData = function(a, b) {
            var e = a[c.expando],
                f = e && d[e];
            f && (b ? c.each(b, function(a, b) {
                delete f[b]
            }) : delete d[e])
        }, c.extend = function() {
            var a, b, d, e, f, g, h = arguments[0] || {},
                i = 1,
                j = arguments.length,
                k = !1;
            for ("boolean" == typeof h && (k = h, h = arguments[i] || {}, i++), "object" != typeof h && "function" !== c.type(h) && (h = {}), i === j && (h = this, i--); i < j; i++)
                if (f = arguments[i])
                    for (e in f) f.hasOwnProperty(e) && (a = h[e], d = f[e], h !== d && (k && d && (c.isPlainObject(d) || (b = c.isArray(d))) ? (b ? (b = !1, g = a && c.isArray(a) ? a : []) : g = a && c.isPlainObject(a) ? a : {}, h[e] = c.extend(k, g, d)) : d !== undefined && (h[e] = d)));
            return h
        }, c.queue = function(a, d, e) {
            if (a) {
                d = (d || "fx") + "queue";
                var f = c.data(a, d);
                return e ? (!f || c.isArray(e) ? f = c.data(a, d, function(a, c) {
                    var d = c || [];
                    return a && (b(Object(a)) ? function(a, b) {
                        for (var c = +b.length, d = 0, e = a.length; d < c;) a[e++] = b[d++];
                        if (c !== c)
                            for (; b[d] !== undefined;) a[e++] = b[d++];
                        a.length = e, a
                    }(d, "string" == typeof a ? [a] : a) : [].push.call(d, a)), d
                }(e)) : f.push(e), f) : f || []
            }
        }, c.dequeue = function(a, b) {
            c.each(a.nodeType ? [a] : a, function(a, d) {
                b = b || "fx";
                var e = c.queue(d, b),
                    f = e.shift();
                "inprogress" === f && (f = e.shift()), f && ("fx" === b && e.unshift("inprogress"), f.call(d, function() {
                    c.dequeue(d, b)
                }))
            })
        }, c.fn = c.prototype = {
            init: function(a) {
                if (a.nodeType) return this[0] = a, this;
                throw new Error("Not a DOM node.")
            },
            offset: function() {
                var b = this[0].getBoundingClientRect ? this[0].getBoundingClientRect() : {
                    top: 0,
                    left: 0
                };
                return {
                    top: b.top + (a.pageYOffset || document.scrollTop || 0) - (document.clientTop || 0),
                    left: b.left + (a.pageXOffset || document.scrollLeft || 0) - (document.clientLeft || 0)
                }
            },
            position: function() {
                var a = this[0],
                    b = function(a) {
                        for (var b = a.offsetParent; b && "html" !== b.nodeName.toLowerCase() && b.style && "static" === b.style.position;) b = b.offsetParent;
                        return b || document
                    }(a),
                    d = this.offset(),
                    e = /^(?:body|html)$/i.test(b.nodeName) ? {
                        top: 0,
                        left: 0
                    } : c(b).offset();
                return d.top -= parseFloat(a.style.marginTop) || 0, d.left -= parseFloat(a.style.marginLeft) || 0, b.style && (e.top += parseFloat(b.style.borderTopWidth) || 0, e.left += parseFloat(b.style.borderLeftWidth) || 0), {
                    top: d.top - e.top,
                    left: d.left - e.left
                }
            }
        };
        var d = {};
        c.expando = "velocity" + (new Date).getTime(), c.uuid = 0;
        for (var e = {}, f = e.hasOwnProperty, g = e.toString, h = "Boolean Number String Function Array Date RegExp Object Error".split(" "), i = 0; i < h.length; i++) e["[object " + h[i] + "]"] = h[i].toLowerCase();
        c.fn.init.prototype = c.fn, a.Velocity = {
            Utilities: c
        }
    }
}(window),
function(a) {
    "use strict";
    "object" == typeof module && "object" == typeof module.exports ? module.exports = a() : "function" == typeof define && define.amd ? define(a) : a()
}(function() {
    "use strict";
    return function(a, b, c, d) {
        function e(a) {
            for (var b = -1, c = a ? a.length : 0, d = []; ++b < c;) {
                var e = a[b];
                e && d.push(e)
            }
            return d
        }

        function f(a) {
            return u.isWrapped(a) ? a = s.call(a) : u.isNode(a) && (a = [a]), a
        }

        function g(a) {
            var b = o.data(a, "velocity");
            return null === b ? d : b
        }

        function h(a, b) {
            var c = g(a);
            c && c.delayTimer && !c.delayPaused && (c.delayRemaining = c.delay - b + c.delayBegin, c.delayPaused = !0, clearTimeout(c.delayTimer.setTimeout))
        }

        function i(a, b) {
            var c = g(a);
            c && c.delayTimer && c.delayPaused && (c.delayPaused = !1, c.delayTimer.setTimeout = setTimeout(c.delayTimer.next, c.delayRemaining))
        }

        function j(a) {
            return function(b) {
                return Math.round(b * a) * (1 / a)
            }
        }

        function k(a, c, d, e) {
            function f(a, b) {
                return 1 - 3 * b + 3 * a
            }

            function g(a, b) {
                return 3 * b - 6 * a
            }

            function h(a) {
                return 3 * a
            }

            function i(a, b, c) {
                return ((f(b, c) * a + g(b, c)) * a + h(b)) * a
            }

            function j(a, b, c) {
                return 3 * f(b, c) * a * a + 2 * g(b, c) * a + h(b)
            }

            function k(b, c) {
                for (var e = 0; e < p; ++e) {
                    var f = j(c, a, d);
                    if (0 === f) return c;
                    c -= (i(c, a, d) - b) / f
                }
                return c
            }

            function l() {
                for (var b = 0; b < t; ++b) x[b] = i(b * u, a, d)
            }

            function m(b, c, e) {
                var f, g, h = 0;
                do {
                    g = c + (e - c) / 2, f = i(g, a, d) - b, f > 0 ? e = g : c = g
                } while (Math.abs(f) > r && ++h < s);
                return g
            }

            function n(b) {
                for (var c = 0, e = 1, f = t - 1; e !== f && x[e] <= b; ++e) c += u;
                --e;
                var g = (b - x[e]) / (x[e + 1] - x[e]),
                    h = c + g * u,
                    i = j(h, a, d);
                return i >= q ? k(b, h) : 0 === i ? h : m(b, c, c + u)
            }

            function o() {
                y = !0, a === c && d === e || l()
            }
            var p = 4,
                q = .001,
                r = 1e-7,
                s = 10,
                t = 11,
                u = 1 / (t - 1),
                v = "Float32Array" in b;
            if (4 !== arguments.length) return !1;
            for (var w = 0; w < 4; ++w)
                if ("number" != typeof arguments[w] || isNaN(arguments[w]) || !isFinite(arguments[w])) return !1;
            a = Math.min(a, 1), d = Math.min(d, 1), a = Math.max(a, 0), d = Math.max(d, 0);
            var x = v ? new Float32Array(t) : new Array(t),
                y = !1,
                z = function(b) {
                    return y || o(), a === c && d === e ? b : 0 === b ? 0 : 1 === b ? 1 : i(n(b), c, e)
                };
            z.getControlPoints = function() {
                return [{
                    x: a,
                    y: c
                }, {
                    x: d,
                    y: e
                }]
            };
            var A = "generateBezier(" + [a, c, d, e] + ")";
            return z.toString = function() {
                return A
            }, z
        }

        function l(a, b) {
            var c = a;
            return u.isString(a) ? y.Easings[a] || (c = !1) : c = u.isArray(a) && 1 === a.length ? j.apply(null, a) : u.isArray(a) && 2 === a.length ? z.apply(null, a.concat([b])) : !(!u.isArray(a) || 4 !== a.length) && k.apply(null, a), c === !1 && (c = y.Easings[y.defaults.easing] ? y.defaults.easing : x), c
        }

        function m(a) {
            if (a) {
                var b = y.timestamp && a !== !0 ? a : r.now(),
                    c = y.State.calls.length;
                c > 1e4 && (y.State.calls = e(y.State.calls), c = y.State.calls.length);
                for (var f = 0; f < c; f++)
                    if (y.State.calls[f]) {
                        var h = y.State.calls[f],
                            i = h[0],
                            j = h[2],
                            k = h[3],
                            l = !!k,
                            q = null,
                            s = h[5],
                            t = h[6];
                        if (k || (k = y.State.calls[f][3] = b - 16), s) {
                            if (s.resume !== !0) continue;
                            k = h[3] = Math.round(b - t - 16), h[5] = null
                        }
                        t = h[6] = b - k;
                        for (var v = Math.min(t / j.duration, 1), w = 0, x = i.length; w < x; w++) {
                            var z = i[w],
                                B = z.element;
                            if (g(B)) {
                                var D = !1;
                                if (j.display !== d && null !== j.display && "none" !== j.display) {
                                    if ("flex" === j.display) {
                                        var E = ["-webkit-box", "-moz-box", "-ms-flexbox", "-webkit-flex"];
                                        o.each(E, function(a, b) {
                                            A.setPropertyValue(B, "display", b)
                                        })
                                    }
                                    A.setPropertyValue(B, "display", j.display)
                                }
                                j.visibility !== d && "hidden" !== j.visibility && A.setPropertyValue(B, "visibility", j.visibility);
                                for (var F in z)
                                    if (z.hasOwnProperty(F) && "element" !== F) {
                                        var G, H = z[F],
                                            I = u.isString(H.easing) ? y.Easings[H.easing] : H.easing;
                                        if (u.isString(H.pattern)) {
                                            var J = 1 === v ? function(a, b, c) {
                                                var d = H.endValue[b];
                                                return c ? Math.round(d) : d
                                            } : function(a, b, c) {
                                                var d = H.startValue[b],
                                                    e = H.endValue[b] - d,
                                                    f = d + e * I(v, j, e);
                                                return c ? Math.round(f) : f
                                            };
                                            G = H.pattern.replace(/{(\d+)(!)?}/g, J)
                                        } else if (1 === v) G = H.endValue;
                                        else {
                                            var K = H.endValue - H.startValue;
                                            G = H.startValue + K * I(v, j, K)
                                        }
                                        if (!l && G === H.currentValue) continue;
                                        if (H.currentValue = G, "tween" === F) q = G;
                                        else {
                                            var L;
                                            if (A.Hooks.registered[F]) {
                                                L = A.Hooks.getRoot(F);
                                                var M = g(B).rootPropertyValueCache[L];
                                                M && (H.rootPropertyValue = M)
                                            }
                                            var N = A.setPropertyValue(B, F, H.currentValue + (p < 9 && 0 === parseFloat(G) ? "" : H.unitType), H.rootPropertyValue, H.scrollData);
                                            A.Hooks.registered[F] && (A.Normalizations.registered[L] ? g(B).rootPropertyValueCache[L] = A.Normalizations.registered[L]("extract", null, N[1]) : g(B).rootPropertyValueCache[L] = N[1]), "transform" === N[0] && (D = !0)
                                        }
                                    }
                                j.mobileHA && g(B).transformCache.translate3d === d && (g(B).transformCache.translate3d = "(0px, 0px, 0px)", D = !0), D && A.flushTransformCache(B)
                            }
                        }
                        j.display !== d && "none" !== j.display && (y.State.calls[f][2].display = !1), j.visibility !== d && "hidden" !== j.visibility && (y.State.calls[f][2].visibility = !1), j.progress && j.progress.call(h[1], h[1], v, Math.max(0, k + j.duration - b), k, q), 1 === v && n(f)
                    }
            }
            y.State.isTicking && C(m)
        }

        function n(a, b) {
            if (!y.State.calls[a]) return !1;
            for (var c = y.State.calls[a][0], e = y.State.calls[a][1], f = y.State.calls[a][2], h = y.State.calls[a][4], i = !1, j = 0, k = c.length; j < k; j++) {
                var l = c[j].element;
                b || f.loop || ("none" === f.display && A.setPropertyValue(l, "display", f.display), "hidden" === f.visibility && A.setPropertyValue(l, "visibility", f.visibility));
                var m = g(l);
                if (f.loop !== !0 && (o.queue(l)[1] === d || !/\.velocityQueueEntryFlag/i.test(o.queue(l)[1])) && m) {
                    m.isAnimating = !1, m.rootPropertyValueCache = {};
                    var n = !1;
                    o.each(A.Lists.transforms3D, function(a, b) {
                        var c = /^scale/.test(b) ? 1 : 0,
                            e = m.transformCache[b];
                        m.transformCache[b] !== d && new RegExp("^\\(" + c + "[^.]").test(e) && (n = !0, delete m.transformCache[b])
                    }), f.mobileHA && (n = !0, delete m.transformCache.translate3d), n && A.flushTransformCache(l), A.Values.removeClass(l, "velocity-animating")
                }
                if (!b && f.complete && !f.loop && j === k - 1) try {
                    f.complete.call(e, e)
                } catch (r) {
                    setTimeout(function() {
                        throw r
                    }, 1)
                }
                h && f.loop !== !0 && h(e), m && f.loop === !0 && !b && (o.each(m.tweensContainer, function(a, b) {
                    if (/^rotate/.test(a) && (parseFloat(b.startValue) - parseFloat(b.endValue)) % 360 == 0) {
                        var c = b.startValue;
                        b.startValue = b.endValue, b.endValue = c
                    }
                    /^backgroundPosition/.test(a) && 100 === parseFloat(b.endValue) && "%" === b.unitType && (b.endValue = 0, b.startValue = 100)
                }), y(l, "reverse", {
                    loop: !0,
                    delay: f.delay
                })), f.queue !== !1 && o.dequeue(l, f.queue)
            }
            y.State.calls[a] = !1;
            for (var p = 0, q = y.State.calls.length; p < q; p++)
                if (y.State.calls[p] !== !1) {
                    i = !0;
                    break
                }
            i === !1 && (y.State.isTicking = !1, delete y.State.calls, y.State.calls = [])
        }
        var o, p = function() {
                if (c.documentMode) return c.documentMode;
                for (var a = 7; a > 4; a--) {
                    var b = c.createElement("div");
                    if (b.innerHTML = "<!--[if IE " + a + "]><span></span><![endif]-->", b.getElementsByTagName("span").length) return b = null, a
                }
                return d
            }(),
            q = function() {
                var a = 0;
                return b.webkitRequestAnimationFrame || b.mozRequestAnimationFrame || function(b) {
                    var c, d = (new Date).getTime();
                    return c = Math.max(0, 16 - (d - a)), a = d + c, setTimeout(function() {
                        b(d + c)
                    }, c)
                }
            }(),
            r = function() {
                var a = b.performance || {};
                if ("function" != typeof a.now) {
                    var c = a.timing && a.timing.navigationStart ? a.timing.navigationStart : (new Date).getTime();
                    a.now = function() {
                        return (new Date).getTime() - c
                    }
                }
                return a
            }(),
            s = function() {
                var a = Array.prototype.slice;
                try {
                    return a.call(c.documentElement), a
                } catch (b) {
                    return function(b, c) {
                        var d = this.length;
                        if ("number" != typeof b && (b = 0), "number" != typeof c && (c = d), this.slice) return a.call(this, b, c);
                        var e, f = [],
                            g = b >= 0 ? b : Math.max(0, d + b),
                            h = c < 0 ? d + c : Math.min(c, d),
                            i = h - g;
                        if (i > 0)
                            if (f = new Array(i), this.charAt)
                                for (e = 0; e < i; e++) f[e] = this.charAt(g + e);
                            else
                                for (e = 0; e < i; e++) f[e] = this[g + e];
                        return f
                    }
                }
            }(),
            t = function() {
                return Array.prototype.includes ? function(a, b) {
                    return a.includes(b)
                } : Array.prototype.indexOf ? function(a, b) {
                    return a.indexOf(b) >= 0
                } : function(a, b) {
                    for (var c = 0; c < a.length; c++)
                        if (a[c] === b) return !0;
                    return !1
                }
            },
            u = {
                isNumber: function(a) {
                    return "number" == typeof a
                },
                isString: function(a) {
                    return "string" == typeof a
                },
                isArray: Array.isArray || function(a) {
                    return "[object Array]" === Object.prototype.toString.call(a)
                },
                isFunction: function(a) {
                    return "[object Function]" === Object.prototype.toString.call(a)
                },
                isNode: function(a) {
                    return a && a.nodeType
                },
                isWrapped: function(a) {
                    return a && a !== b && u.isNumber(a.length) && !u.isString(a) && !u.isFunction(a) && !u.isNode(a) && (0 === a.length || u.isNode(a[0]))
                },
                isSVG: function(a) {
                    return b.SVGElement && a instanceof b.SVGElement
                },
                isEmptyObject: function(a) {
                    for (var b in a)
                        if (a.hasOwnProperty(b)) return !1;
                    return !0
                }
            },
            v = !1;
        if (a.fn && a.fn.jquery ? (o = a, v = !0) : o = b.Velocity.Utilities, p <= 8 && !v) throw new Error("Velocity: IE8 and below require jQuery to be loaded before Velocity.");
        if (p <= 7) return void(jQuery.fn.velocity = jQuery.fn.animate);
        var w = 400,
            x = "swing",
            y = {
                State: {
                    isMobile: /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent),
                    isAndroid: /Android/i.test(navigator.userAgent),
                    isGingerbread: /Android 2\.3\.[3-7]/i.test(navigator.userAgent),
                    isChrome: b.chrome,
                    isFirefox: /Firefox/i.test(navigator.userAgent),
                    prefixElement: c.createElement("div"),
                    prefixMatches: {},
                    scrollAnchor: null,
                    scrollPropertyLeft: null,
                    scrollPropertyTop: null,
                    isTicking: !1,
                    calls: [],
                    delayedElements: {
                        count: 0
                    }
                },
                CSS: {},
                Utilities: o,
                Redirects: {},
                Easings: {},
                Promise: b.Promise,
                defaults: {
                    queue: "",
                    duration: w,
                    easing: x,
                    begin: d,
                    complete: d,
                    progress: d,
                    display: d,
                    visibility: d,
                    loop: !1,
                    delay: !1,
                    mobileHA: !0,
                    _cacheValues: !0,
                    promiseRejectEmpty: !0
                },
                init: function(a) {
                    o.data(a, "velocity", {
                        isSVG: u.isSVG(a),
                        isAnimating: !1,
                        computedStyle: null,
                        tweensContainer: null,
                        rootPropertyValueCache: {},
                        transformCache: {}
                    })
                },
                hook: null,
                mock: !1,
                version: {
                    major: 1,
                    minor: 5,
                    patch: 0
                },
                debug: !1,
                timestamp: !0,
                pauseAll: function(a) {
                    var b = (new Date).getTime();
                    o.each(y.State.calls, function(b, c) {
                        if (c) {
                            if (a !== d && (c[2].queue !== a || c[2].queue === !1)) return !0;
                            c[5] = {
                                resume: !1
                            }
                        }
                    }), o.each(y.State.delayedElements, function(a, c) {
                        c && h(c, b)
                    })
                },
                resumeAll: function(a) {
                    var b = (new Date).getTime();
                    o.each(y.State.calls, function(b, c) {
                        if (c) {
                            if (a !== d && (c[2].queue !== a || c[2].queue === !1)) return !0;
                            c[5] && (c[5].resume = !0)
                        }
                    }), o.each(y.State.delayedElements, function(a, c) {
                        c && i(c, b)
                    })
                }
            };
        b.pageYOffset !== d ? (y.State.scrollAnchor = b, y.State.scrollPropertyLeft = "pageXOffset", y.State.scrollPropertyTop = "pageYOffset") : (y.State.scrollAnchor = c.documentElement || c.body.parentNode || c.body, y.State.scrollPropertyLeft = "scrollLeft", y.State.scrollPropertyTop = "scrollTop");
        var z = function() {
            function a(a) {
                return -a.tension * a.x - a.friction * a.v
            }

            function b(b, c, d) {
                var e = {
                    x: b.x + d.dx * c,
                    v: b.v + d.dv * c,
                    tension: b.tension,
                    friction: b.friction
                };
                return {
                    dx: e.v,
                    dv: a(e)
                }
            }

            function c(c, d) {
                var e = {
                        dx: c.v,
                        dv: a(c)
                    },
                    f = b(c, .5 * d, e),
                    g = b(c, .5 * d, f),
                    h = b(c, d, g),
                    i = 1 / 6 * (e.dx + 2 * (f.dx + g.dx) + h.dx),
                    j = 1 / 6 * (e.dv + 2 * (f.dv + g.dv) + h.dv);
                return c.x = c.x + i * d, c.v = c.v + j * d, c
            }
            return function d(a, b, e) {
                var f, g, h, i = {
                        x: -1,
                        v: 0,
                        tension: null,
                        friction: null
                    },
                    j = [0],
                    k = 0;
                for (a = parseFloat(a) || 500, b = parseFloat(b) || 20, e = e || null, i.tension = a, i.friction = b, f = null !== e, f ? (k = d(a, b), g = k / e * .016) : g = .016;;)
                    if (h = c(h || i, g), j.push(1 + h.x), k += 16, !(Math.abs(h.x) > 1e-4 && Math.abs(h.v) > 1e-4)) break;
                return f ? function(a) {
                    return j[a * (j.length - 1) | 0]
                } : k
            }
        }();
        y.Easings = {
            linear: function(a) {
                return a
            },
            swing: function(a) {
                return .5 - Math.cos(a * Math.PI) / 2
            },
            spring: function(a) {
                return 1 - Math.cos(4.5 * a * Math.PI) * Math.exp(6 * -a)
            }
        }, o.each([
            ["ease", [.25, .1, .25, 1]],
            ["ease-in", [.42, 0, 1, 1]],
            ["ease-out", [0, 0, .58, 1]],
            ["ease-in-out", [.42, 0, .58, 1]],
            ["easeInSine", [.47, 0, .745, .715]],
            ["easeOutSine", [.39, .575, .565, 1]],
            ["easeInOutSine", [.445, .05, .55, .95]],
            ["easeInQuad", [.55, .085, .68, .53]],
            ["easeOutQuad", [.25, .46, .45, .94]],
            ["easeInOutQuad", [.455, .03, .515, .955]],
            ["easeInCubic", [.55, .055, .675, .19]],
            ["easeOutCubic", [.215, .61, .355, 1]],
            ["easeInOutCubic", [.645, .045, .355, 1]],
            ["easeInQuart", [.895, .03, .685, .22]],
            ["easeOutQuart", [.165, .84, .44, 1]],
            ["easeInOutQuart", [.77, 0, .175, 1]],
            ["easeInQuint", [.755, .05, .855, .06]],
            ["easeOutQuint", [.23, 1, .32, 1]],
            ["easeInOutQuint", [.86, 0, .07, 1]],
            ["easeInExpo", [.95, .05, .795, .035]],
            ["easeOutExpo", [.19, 1, .22, 1]],
            ["easeInOutExpo", [1, 0, 0, 1]],
            ["easeInCirc", [.6, .04, .98, .335]],
            ["easeOutCirc", [.075, .82, .165, 1]],
            ["easeInOutCirc", [.785, .135, .15, .86]]
        ], function(a, b) {
            y.Easings[b[0]] = k.apply(null, b[1])
        });
        var A = y.CSS = {
            RegEx: {
                isHex: /^#([A-f\d]{3}){1,2}$/i,
                valueUnwrap: /^[A-z]+\((.*)\)$/i,
                wrappedValueAlreadyExtracted: /[0-9.]+ [0-9.]+ [0-9.]+( [0-9.]+)?/,
                valueSplit: /([A-z]+\(.+\))|(([A-z0-9#-.]+?)(?=\s|$))/gi
            },
            Lists: {
                colors: ["fill", "stroke", "stopColor", "color", "backgroundColor", "borderColor", "borderTopColor", "borderRightColor", "borderBottomColor", "borderLeftColor", "outlineColor"],
                transformsBase: ["translateX", "translateY", "scale", "scaleX", "scaleY", "skewX", "skewY", "rotateZ"],
                transforms3D: ["transformPerspective", "translateZ", "scaleZ", "rotateX", "rotateY"],
                units: ["%", "em", "ex", "ch", "rem", "vw", "vh", "vmin", "vmax", "cm", "mm", "Q", "in", "pc", "pt", "px", "deg", "grad", "rad", "turn", "s", "ms"],
                colorNames: {
                    aliceblue: "240,248,255",
                    antiquewhite: "250,235,215",
                    aquamarine: "127,255,212",
                    aqua: "0,255,255",
                    azure: "240,255,255",
                    beige: "245,245,220",
                    bisque: "255,228,196",
                    black: "0,0,0",
                    blanchedalmond: "255,235,205",
                    blueviolet: "138,43,226",
                    blue: "0,0,255",
                    brown: "165,42,42",
                    burlywood: "222,184,135",
                    cadetblue: "95,158,160",
                    chartreuse: "127,255,0",
                    chocolate: "210,105,30",
                    coral: "255,127,80",
                    cornflowerblue: "100,149,237",
                    cornsilk: "255,248,220",
                    crimson: "220,20,60",
                    cyan: "0,255,255",
                    darkblue: "0,0,139",
                    darkcyan: "0,139,139",
                    darkgoldenrod: "184,134,11",
                    darkgray: "169,169,169",
                    darkgrey: "169,169,169",
                    darkgreen: "0,100,0",
                    darkkhaki: "189,183,107",
                    darkmagenta: "139,0,139",
                    darkolivegreen: "85,107,47",
                    darkorange: "255,140,0",
                    darkorchid: "153,50,204",
                    darkred: "139,0,0",
                    darksalmon: "233,150,122",
                    darkseagreen: "143,188,143",
                    darkslateblue: "72,61,139",
                    darkslategray: "47,79,79",
                    darkturquoise: "0,206,209",
                    darkviolet: "148,0,211",
                    deeppink: "255,20,147",
                    deepskyblue: "0,191,255",
                    dimgray: "105,105,105",
                    dimgrey: "105,105,105",
                    dodgerblue: "30,144,255",
                    firebrick: "178,34,34",
                    floralwhite: "255,250,240",
                    forestgreen: "34,139,34",
                    fuchsia: "255,0,255",
                    gainsboro: "220,220,220",
                    ghostwhite: "248,248,255",
                    gold: "255,215,0",
                    goldenrod: "218,165,32",
                    gray: "128,128,128",
                    grey: "128,128,128",
                    greenyellow: "173,255,47",
                    green: "0,128,0",
                    honeydew: "240,255,240",
                    hotpink: "255,105,180",
                    indianred: "205,92,92",
                    indigo: "75,0,130",
                    ivory: "255,255,240",
                    khaki: "240,230,140",
                    lavenderblush: "255,240,245",
                    lavender: "230,230,250",
                    lawngreen: "124,252,0",
                    lemonchiffon: "255,250,205",
                    lightblue: "173,216,230",
                    lightcoral: "240,128,128",
                    lightcyan: "224,255,255",
                    lightgoldenrodyellow: "250,250,210",
                    lightgray: "211,211,211",
                    lightgrey: "211,211,211",
                    lightgreen: "144,238,144",
                    lightpink: "255,182,193",
                    lightsalmon: "255,160,122",
                    lightseagreen: "32,178,170",
                    lightskyblue: "135,206,250",
                    lightslategray: "119,136,153",
                    lightsteelblue: "176,196,222",
                    lightyellow: "255,255,224",
                    limegreen: "50,205,50",
                    lime: "0,255,0",
                    linen: "250,240,230",
                    magenta: "255,0,255",
                    maroon: "128,0,0",
                    mediumaquamarine: "102,205,170",
                    mediumblue: "0,0,205",
                    mediumorchid: "186,85,211",
                    mediumpurple: "147,112,219",
                    mediumseagreen: "60,179,113",
                    mediumslateblue: "123,104,238",
                    mediumspringgreen: "0,250,154",
                    mediumturquoise: "72,209,204",
                    mediumvioletred: "199,21,133",
                    midnightblue: "25,25,112",
                    mintcream: "245,255,250",
                    mistyrose: "255,228,225",
                    moccasin: "255,228,181",
                    navajowhite: "255,222,173",
                    navy: "0,0,128",
                    oldlace: "253,245,230",
                    olivedrab: "107,142,35",
                    olive: "128,128,0",
                    orangered: "255,69,0",
                    orange: "255,165,0",
                    orchid: "218,112,214",
                    palegoldenrod: "238,232,170",
                    palegreen: "152,251,152",
                    paleturquoise: "175,238,238",
                    palevioletred: "219,112,147",
                    papayawhip: "255,239,213",
                    peachpuff: "255,218,185",
                    peru: "205,133,63",
                    pink: "255,192,203",
                    plum: "221,160,221",
                    powderblue: "176,224,230",
                    purple: "128,0,128",
                    red: "255,0,0",
                    rosybrown: "188,143,143",
                    royalblue: "65,105,225",
                    saddlebrown: "139,69,19",
                    salmon: "250,128,114",
                    sandybrown: "244,164,96",
                    seagreen: "46,139,87",
                    seashell: "255,245,238",
                    sienna: "160,82,45",
                    silver: "192,192,192",
                    skyblue: "135,206,235",
                    slateblue: "106,90,205",
                    slategray: "112,128,144",
                    snow: "255,250,250",
                    springgreen: "0,255,127",
                    steelblue: "70,130,180",
                    tan: "210,180,140",
                    teal: "0,128,128",
                    thistle: "216,191,216",
                    tomato: "255,99,71",
                    turquoise: "64,224,208",
                    violet: "238,130,238",
                    wheat: "245,222,179",
                    whitesmoke: "245,245,245",
                    white: "255,255,255",
                    yellowgreen: "154,205,50",
                    yellow: "255,255,0"
                }
            },
            Hooks: {
                templates: {
                    textShadow: ["Color X Y Blur", "black 0px 0px 0px"],
                    boxShadow: ["Color X Y Blur Spread", "black 0px 0px 0px 0px"],
                    clip: ["Top Right Bottom Left", "0px 0px 0px 0px"],
                    backgroundPosition: ["X Y", "0% 0%"],
                    transformOrigin: ["X Y Z", "50% 50% 0px"],
                    perspectiveOrigin: ["X Y", "50% 50%"]
                },
                registered: {},
                register: function() {
                    for (var a = 0; a < A.Lists.colors.length; a++) {
                        var b = "color" === A.Lists.colors[a] ? "0 0 0 1" : "255 255 255 1";
                        A.Hooks.templates[A.Lists.colors[a]] = ["Red Green Blue Alpha", b]
                    }
                    var c, d, e;
                    if (p)
                        for (c in A.Hooks.templates)
                            if (A.Hooks.templates.hasOwnProperty(c)) {
                                d = A.Hooks.templates[c], e = d[0].split(" ");
                                var f = d[1].match(A.RegEx.valueSplit);
                                "Color" === e[0] && (e.push(e.shift()), f.push(f.shift()), A.Hooks.templates[c] = [e.join(" "), f.join(" ")])
                            }
                    for (c in A.Hooks.templates)
                        if (A.Hooks.templates.hasOwnProperty(c)) {
                            d = A.Hooks.templates[c], e = d[0].split(" ");
                            for (var g in e)
                                if (e.hasOwnProperty(g)) {
                                    var h = c + e[g],
                                        i = g;
                                    A.Hooks.registered[h] = [c, i]
                                }
                        }
                },
                getRoot: function(a) {
                    var b = A.Hooks.registered[a];
                    return b ? b[0] : a
                },
                getUnit: function(a, b) {
                    var c = (a.substr(b || 0, 5).match(/^[a-z%]+/) || [])[0] || "";
                    return c && t(A.Lists.units, c) ? c : ""
                },
                fixColors: function(a) {
                    return a.replace(/(rgba?\(\s*)?(\b[a-z]+\b)/g, function(a, b, c) {
                        return A.Lists.colorNames.hasOwnProperty(c) ? (b ? b : "rgba(") + A.Lists.colorNames[c] + (b ? "" : ",1)") : b + c
                    })
                },
                cleanRootPropertyValue: function(a, b) {
                    return A.RegEx.valueUnwrap.test(b) && (b = b.match(A.RegEx.valueUnwrap)[1]), A.Values.isCSSNullValue(b) && (b = A.Hooks.templates[a][1]), b
                },
                extractValue: function(a, b) {
                    var c = A.Hooks.registered[a];
                    if (c) {
                        var d = c[0],
                            e = c[1];
                        return b = A.Hooks.cleanRootPropertyValue(d, b), b.toString().match(A.RegEx.valueSplit)[e]
                    }
                    return b
                },
                injectValue: function(a, b, c) {
                    var d = A.Hooks.registered[a];
                    if (d) {
                        var e, f = d[0],
                            g = d[1];
                        return c = A.Hooks.cleanRootPropertyValue(f, c), e = c.toString().match(A.RegEx.valueSplit), e[g] = b, e.join(" ")
                    }
                    return c
                }
            },
            Normalizations: {
                registered: {
                    clip: function(a, b, c) {
                        switch (a) {
                            case "name":
                                return "clip";
                            case "extract":
                                var d;
                                return A.RegEx.wrappedValueAlreadyExtracted.test(c) ? d = c : (d = c.toString().match(A.RegEx.valueUnwrap), d = d ? d[1].replace(/,(\s+)?/g, " ") : c), d;
                            case "inject":
                                return "rect(" + c + ")"
                        }
                    },
                    blur: function(a, b, c) {
                        switch (a) {
                            case "name":
                                return y.State.isFirefox ? "filter" : "-webkit-filter";
                            case "extract":
                                var d = parseFloat(c);
                                if (!d && 0 !== d) {
                                    var e = c.toString().match(/blur\(([0-9]+[A-z]+)\)/i);
                                    d = e ? e[1] : 0
                                }
                                return d;
                            case "inject":
                                return parseFloat(c) ? "blur(" + c + ")" : "none"
                        }
                    },
                    opacity: function(a, b, c) {
                        if (p <= 8) switch (a) {
                            case "name":
                                return "filter";
                            case "extract":
                                var d = c.toString().match(/alpha\(opacity=(.*)\)/i);
                                return c = d ? d[1] / 100 : 1;
                            case "inject":
                                return b.style.zoom = 1, parseFloat(c) >= 1 ? "" : "alpha(opacity=" + parseInt(100 * parseFloat(c), 10) + ")"
                        } else switch (a) {
                            case "name":
                                return "opacity";
                            case "extract":
                                return c;
                            case "inject":
                                return c
                        }
                    }
                },
                register: function() {
                    function a(a, b, c) {
                        if ("border-box" === A.getPropertyValue(b, "boxSizing").toString().toLowerCase() === (c || !1)) {
                            var d, e, f = 0,
                                g = "width" === a ? ["Left", "Right"] : ["Top", "Bottom"],
                                h = ["padding" + g[0], "padding" + g[1], "border" + g[0] + "Width", "border" + g[1] + "Width"];
                            for (d = 0; d < h.length; d++) e = parseFloat(A.getPropertyValue(b, h[d])), isNaN(e) || (f += e);
                            return c ? -f : f
                        }
                        return 0
                    }

                    function b(b, c) {
                        return function(d, e, f) {
                            switch (d) {
                                case "name":
                                    return b;
                                case "extract":
                                    return parseFloat(f) + a(b, e, c);
                                case "inject":
                                    return parseFloat(f) - a(b, e, c) + "px"
                            }
                        }
                    }
                    p && !(p > 9) || y.State.isGingerbread || (A.Lists.transformsBase = A.Lists.transformsBase.concat(A.Lists.transforms3D));
                    for (var c = 0; c < A.Lists.transformsBase.length; c++) ! function() {
                        var a = A.Lists.transformsBase[c];
                        A.Normalizations.registered[a] = function(b, c, e) {
                            switch (b) {
                                case "name":
                                    return "transform";
                                case "extract":
                                    return g(c) === d || g(c).transformCache[a] === d ? /^scale/i.test(a) ? 1 : 0 : g(c).transformCache[a].replace(/[()]/g, "");
                                case "inject":
                                    var f = !1;
                                    switch (a.substr(0, a.length - 1)) {
                                        case "translate":
                                            f = !/(%|px|em|rem|vw|vh|\d)$/i.test(e);
                                            break;
                                        case "scal":
                                        case "scale":
                                            y.State.isAndroid && g(c).transformCache[a] === d && e < 1 && (e = 1), f = !/(\d)$/i.test(e);
                                            break;
                                        case "skew":
                                            f = !/(deg|\d)$/i.test(e);
                                            break;
                                        case "rotate":
                                            f = !/(deg|\d)$/i.test(e)
                                    }
                                    return f || (g(c).transformCache[a] = "(" + e + ")"), g(c).transformCache[a]
                            }
                        }
                    }();
                    for (var e = 0; e < A.Lists.colors.length; e++) ! function() {
                        var a = A.Lists.colors[e];
                        A.Normalizations.registered[a] = function(b, c, e) {
                            switch (b) {
                                case "name":
                                    return a;
                                case "extract":
                                    var f;
                                    if (A.RegEx.wrappedValueAlreadyExtracted.test(e)) f = e;
                                    else {
                                        var g, h = {
                                            black: "rgb(0, 0, 0)",
                                            blue: "rgb(0, 0, 255)",
                                            gray: "rgb(128, 128, 128)",
                                            green: "rgb(0, 128, 0)",
                                            red: "rgb(255, 0, 0)",
                                            white: "rgb(255, 255, 255)"
                                        };
                                        /^[A-z]+$/i.test(e) ? g = h[e] !== d ? h[e] : h.black : A.RegEx.isHex.test(e) ? g = "rgb(" + A.Values.hexToRgb(e).join(" ") + ")" : /^rgba?\(/i.test(e) || (g = h.black), f = (g || e).toString().match(A.RegEx.valueUnwrap)[1].replace(/,(\s+)?/g, " ")
                                    }
                                    return (!p || p > 8) && 3 === f.split(" ").length && (f += " 1"), f;
                                case "inject":
                                    return /^rgb/.test(e) ? e : (p <= 8 ? 4 === e.split(" ").length && (e = e.split(/\s+/).slice(0, 3).join(" ")) : 3 === e.split(" ").length && (e += " 1"), (p <= 8 ? "rgb" : "rgba") + "(" + e.replace(/\s+/g, ",").replace(/\.(\d)+(?=,)/g, "") + ")")
                            }
                        }
                    }();
                    A.Normalizations.registered.innerWidth = b("width", !0), A.Normalizations.registered.innerHeight = b("height", !0), A.Normalizations.registered.outerWidth = b("width"), A.Normalizations.registered.outerHeight = b("height")
                }
            },
            Names: {
                camelCase: function(a) {
                    return a.replace(/-(\w)/g, function(a, b) {
                        return b.toUpperCase()
                    })
                },
                SVGAttribute: function(a) {
                    var b = "width|height|x|y|cx|cy|r|rx|ry|x1|x2|y1|y2";
                    return (p || y.State.isAndroid && !y.State.isChrome) && (b += "|transform"), new RegExp("^(" + b + ")$", "i").test(a)
                },
                prefixCheck: function(a) {
                    if (y.State.prefixMatches[a]) return [y.State.prefixMatches[a], !0];
                    for (var b = ["", "Webkit", "Moz", "ms", "O"], c = 0, d = b.length; c < d; c++) {
                        var e;
                        if (e = 0 === c ? a : b[c] + a.replace(/^\w/, function(a) {
                                return a.toUpperCase()
                            }), u.isString(y.State.prefixElement.style[e])) return y.State.prefixMatches[a] = e, [e, !0]
                    }
                    return [a, !1]
                }
            },
            Values: {
                hexToRgb: function(a) {
                    var b, c = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i;
                    return a = a.replace(/^#?([a-f\d])([a-f\d])([a-f\d])$/i, function(a, b, c, d) {
                        return b + b + c + c + d + d
                    }), b = c.exec(a), b ? [parseInt(b[1], 16), parseInt(b[2], 16), parseInt(b[3], 16)] : [0, 0, 0]
                },
                isCSSNullValue: function(a) {
                    return !a || /^(none|auto|transparent|(rgba\(0, ?0, ?0, ?0\)))$/i.test(a)
                },
                getUnitType: function(a) {
                    return /^(rotate|skew)/i.test(a) ? "deg" : /(^(scale|scaleX|scaleY|scaleZ|alpha|flexGrow|flexHeight|zIndex|fontWeight)$)|((opacity|red|green|blue|alpha)$)/i.test(a) ? "" : "px"
                },
                getDisplayType: function(a) {
                    var b = a && a.tagName.toString().toLowerCase();
                    return /^(b|big|i|small|tt|abbr|acronym|cite|code|dfn|em|kbd|strong|samp|var|a|bdo|br|img|map|object|q|script|span|sub|sup|button|input|label|select|textarea)$/i.test(b) ? "inline" : /^(li)$/i.test(b) ? "list-item" : /^(tr)$/i.test(b) ? "table-row" : /^(table)$/i.test(b) ? "table" : /^(tbody)$/i.test(b) ? "table-row-group" : "block"
                },
                addClass: function(a, b) {
                    if (a)
                        if (a.classList) a.classList.add(b);
                        else if (u.isString(a.className)) a.className += (a.className.length ? " " : "") + b;
                    else {
                        var c = a.getAttribute(p <= 7 ? "className" : "class") || "";
                        a.setAttribute("class", c + (c ? " " : "") + b)
                    }
                },
                removeClass: function(a, b) {
                    if (a)
                        if (a.classList) a.classList.remove(b);
                        else if (u.isString(a.className)) a.className = a.className.toString().replace(new RegExp("(^|\\s)" + b.split(" ").join("|") + "(\\s|$)", "gi"), " ");
                    else {
                        var c = a.getAttribute(p <= 7 ? "className" : "class") || "";
                        a.setAttribute("class", c.replace(new RegExp("(^|s)" + b.split(" ").join("|") + "(s|$)", "gi"), " "))
                    }
                }
            },
            getPropertyValue: function(a, c, e, f) {
                function h(a, c) {
                    var e = 0;
                    if (p <= 8) e = o.css(a, c);
                    else {
                        var i = !1;
                        /^(width|height)$/.test(c) && 0 === A.getPropertyValue(a, "display") && (i = !0, A.setPropertyValue(a, "display", A.Values.getDisplayType(a)));
                        var j = function() {
                            i && A.setPropertyValue(a, "display", "none")
                        };
                        if (!f) {
                            if ("height" === c && "border-box" !== A.getPropertyValue(a, "boxSizing").toString().toLowerCase()) {
                                var k = a.offsetHeight - (parseFloat(A.getPropertyValue(a, "borderTopWidth")) || 0) - (parseFloat(A.getPropertyValue(a, "borderBottomWidth")) || 0) - (parseFloat(A.getPropertyValue(a, "paddingTop")) || 0) - (parseFloat(A.getPropertyValue(a, "paddingBottom")) || 0);
                                return j(), k
                            }
                            if ("width" === c && "border-box" !== A.getPropertyValue(a, "boxSizing").toString().toLowerCase()) {
                                var l = a.offsetWidth - (parseFloat(A.getPropertyValue(a, "borderLeftWidth")) || 0) - (parseFloat(A.getPropertyValue(a, "borderRightWidth")) || 0) - (parseFloat(A.getPropertyValue(a, "paddingLeft")) || 0) - (parseFloat(A.getPropertyValue(a, "paddingRight")) || 0);
                                return j(), l
                            }
                        }
                        var m;
                        m = g(a) === d ? b.getComputedStyle(a, null) : g(a).computedStyle ? g(a).computedStyle : g(a).computedStyle = b.getComputedStyle(a, null), "borderColor" === c && (c = "borderTopColor"), e = 9 === p && "filter" === c ? m.getPropertyValue(c) : m[c], "" !== e && null !== e || (e = a.style[c]), j()
                    }
                    if ("auto" === e && /^(top|right|bottom|left)$/i.test(c)) {
                        var n = h(a, "position");
                        ("fixed" === n || "absolute" === n && /top|left/i.test(c)) && (e = o(a).position()[c] + "px")
                    }
                    return e
                }
                var i;
                if (A.Hooks.registered[c]) {
                    var j = c,
                        k = A.Hooks.getRoot(j);
                    e === d && (e = A.getPropertyValue(a, A.Names.prefixCheck(k)[0])), A.Normalizations.registered[k] && (e = A.Normalizations.registered[k]("extract", a, e)), i = A.Hooks.extractValue(j, e)
                } else if (A.Normalizations.registered[c]) {
                    var l, m;
                    l = A.Normalizations.registered[c]("name", a), "transform" !== l && (m = h(a, A.Names.prefixCheck(l)[0]), A.Values.isCSSNullValue(m) && A.Hooks.templates[c] && (m = A.Hooks.templates[c][1])), i = A.Normalizations.registered[c]("extract", a, m)
                }
                if (!/^[\d-]/.test(i)) {
                    var n = g(a);
                    if (n && n.isSVG && A.Names.SVGAttribute(c))
                        if (/^(height|width)$/i.test(c)) try {
                            i = a.getBBox()[c]
                        } catch (q) {
                            i = 0
                        } else i = a.getAttribute(c);
                        else i = h(a, A.Names.prefixCheck(c)[0])
                }
                return A.Values.isCSSNullValue(i) && (i = 0), y.debug >= 2 && console.log("Get " + c + ": " + i), i
            },
            setPropertyValue: function(a, c, d, e, f) {
                var h = c;
                if ("scroll" === c) f.container ? f.container["scroll" + f.direction] = d : "Left" === f.direction ? b.scrollTo(d, f.alternateValue) : b.scrollTo(f.alternateValue, d);
                else if (A.Normalizations.registered[c] && "transform" === A.Normalizations.registered[c]("name", a)) A.Normalizations.registered[c]("inject", a, d), h = "transform", d = g(a).transformCache[c];
                else {
                    if (A.Hooks.registered[c]) {
                        var i = c,
                            j = A.Hooks.getRoot(c);
                        e = e || A.getPropertyValue(a, j), d = A.Hooks.injectValue(i, d, e), c = j
                    }
                    if (A.Normalizations.registered[c] && (d = A.Normalizations.registered[c]("inject", a, d), c = A.Normalizations.registered[c]("name", a)), h = A.Names.prefixCheck(c)[0], p <= 8) try {
                        a.style[h] = d
                    } catch (l) {
                        y.debug && console.log("Browser does not support [" + d + "] for [" + h + "]")
                    } else {
                        var k = g(a);
                        k && k.isSVG && A.Names.SVGAttribute(c) ? a.setAttribute(c, d) : a.style[h] = d
                    }
                    y.debug >= 2 && console.log("Set " + c + " (" + h + "): " + d)
                }
                return [h, d]
            },
            flushTransformCache: function(a) {
                var b = "",
                    c = g(a);
                if ((p || y.State.isAndroid && !y.State.isChrome) && c && c.isSVG) {
                    var d = function(b) {
                            return parseFloat(A.getPropertyValue(a, b))
                        },
                        e = {
                            translate: [d("translateX"), d("translateY")],
                            skewX: [d("skewX")],
                            skewY: [d("skewY")],
                            scale: 1 !== d("scale") ? [d("scale"), d("scale")] : [d("scaleX"), d("scaleY")],
                            rotate: [d("rotateZ"), 0, 0]
                        };
                    o.each(g(a).transformCache, function(a) {
                        /^translate/i.test(a) ? a = "translate" : /^scale/i.test(a) ? a = "scale" : /^rotate/i.test(a) && (a = "rotate"), e[a] && (b += a + "(" + e[a].join(" ") + ") ", delete e[a])
                    })
                } else {
                    var f, h;
                    o.each(g(a).transformCache, function(c) {
                        if (f = g(a).transformCache[c], "transformPerspective" === c) return h = f, !0;
                        9 === p && "rotateZ" === c && (c = "rotate"), b += c + f + " "
                    }), h && (b = "perspective" + h + " " + b)
                }
                A.setPropertyValue(a, "transform", b)
            }
        };
        A.Hooks.register(), A.Normalizations.register(), y.hook = function(a, b, c) {
            var e;
            return a = f(a), o.each(a, function(a, f) {
                if (g(f) === d && y.init(f), c === d) e === d && (e = A.getPropertyValue(f, b));
                else {
                    var h = A.setPropertyValue(f, b, c);
                    "transform" === h[0] && y.CSS.flushTransformCache(f), e = h
                }
            }), e
        };
        var B = function() {
            function a() {
                return k ? z.promise || null : p
            }

            function e(a, e) {
                function f(f) {
                    var k, n;
                    if (i.begin && 0 === D) try {
                        i.begin.call(r, r)
                    } catch (V) {
                        setTimeout(function() {
                            throw V
                        }, 1)
                    }
                    if ("scroll" === G) {
                        var p, q, w, x = /^x$/i.test(i.axis) ? "Left" : "Top",
                            B = parseFloat(i.offset) || 0;
                        i.container ? u.isWrapped(i.container) || u.isNode(i.container) ? (i.container = i.container[0] || i.container, p = i.container["scroll" + x], w = p + o(a).position()[x.toLowerCase()] + B) : i.container = null : (p = y.State.scrollAnchor[y.State["scrollProperty" + x]], q = y.State.scrollAnchor[y.State["scrollProperty" + ("Left" === x ? "Top" : "Left")]], w = o(a).offset()[x.toLowerCase()] + B), j = {
                            scroll: {
                                rootPropertyValue: !1,
                                startValue: p,
                                currentValue: p,
                                endValue: w,
                                unitType: "",
                                easing: i.easing,
                                scrollData: {
                                    container: i.container,
                                    direction: x,
                                    alternateValue: q
                                }
                            },
                            element: a
                        }, y.debug && console.log("tweensContainer (scroll): ", j.scroll, a)
                    } else if ("reverse" === G) {
                        if (!(k = g(a))) return;
                        if (!k.tweensContainer) return void o.dequeue(a, i.queue);
                        "none" === k.opts.display && (k.opts.display = "auto"), "hidden" === k.opts.visibility && (k.opts.visibility = "visible"), k.opts.loop = !1, k.opts.begin = null, k.opts.complete = null, v.easing || delete i.easing, v.duration || delete i.duration, i = o.extend({}, k.opts, i), n = o.extend(!0, {}, k ? k.tweensContainer : null);
                        for (var E in n)
                            if (n.hasOwnProperty(E) && "element" !== E) {
                                var F = n[E].startValue;
                                n[E].startValue = n[E].currentValue = n[E].endValue, n[E].endValue = F, u.isEmptyObject(v) || (n[E].easing = i.easing), y.debug && console.log("reverse tweensContainer (" + E + "): " + JSON.stringify(n[E]), a)
                            }
                        j = n
                    } else if ("start" === G) {
                        k = g(a), k && k.tweensContainer && k.isAnimating === !0 && (n = k.tweensContainer);
                        var H = function(e, f) {
                            var g, l = A.Hooks.getRoot(e),
                                m = !1,
                                p = f[0],
                                q = f[1],
                                r = f[2];
                            if (!(k && k.isSVG || "tween" === l || A.Names.prefixCheck(l)[1] !== !1 || A.Normalizations.registered[l] !== d)) return void(y.debug && console.log("Skipping [" + l + "] due to a lack of browser support."));
                            (i.display !== d && null !== i.display && "none" !== i.display || i.visibility !== d && "hidden" !== i.visibility) && /opacity|filter/.test(e) && !r && 0 !== p && (r = 0), i._cacheValues && n && n[e] ? (r === d && (r = n[e].endValue + n[e].unitType), m = k.rootPropertyValueCache[l]) : A.Hooks.registered[e] ? r === d ? (m = A.getPropertyValue(a, l), r = A.getPropertyValue(a, e, m)) : m = A.Hooks.templates[l][1] : r === d && (r = A.getPropertyValue(a, e));
                            var s, t, v, w = !1,
                                x = function(a, b) {
                                    var c, d;
                                    return d = (b || "0").toString().toLowerCase().replace(/[%A-z]+$/, function(a) {
                                        return c = a, ""
                                    }), c || (c = A.Values.getUnitType(a)), [d, c]
                                };
                            if (r !== p && u.isString(r) && u.isString(p)) {
                                g = "";
                                var z = 0,
                                    B = 0,
                                    C = [],
                                    D = [],
                                    E = 0,
                                    F = 0,
                                    G = 0;
                                for (r = A.Hooks.fixColors(r), p = A.Hooks.fixColors(p); z < r.length && B < p.length;) {
                                    var H = r[z],
                                        I = p[B];
                                    if (/[\d\.-]/.test(H) && /[\d\.-]/.test(I)) {
                                        for (var J = H, K = I, L = ".", N = "."; ++z < r.length;) {
                                            if ((H = r[z]) === L) L = "..";
                                            else if (!/\d/.test(H)) break;
                                            J += H
                                        }
                                        for (; ++B < p.length;) {
                                            if ((I = p[B]) === N) N = "..";
                                            else if (!/\d/.test(I)) break;
                                            K += I
                                        }
                                        var O = A.Hooks.getUnit(r, z),
                                            P = A.Hooks.getUnit(p, B);
                                        if (z += O.length, B += P.length, O === P) J === K ? g += J + O : (g += "{" + C.length + (F ? "!" : "") + "}" + O, C.push(parseFloat(J)), D.push(parseFloat(K)));
                                        else {
                                            var Q = parseFloat(J),
                                                R = parseFloat(K);
                                            g += (E < 5 ? "calc" : "") + "(" + (Q ? "{" + C.length + (F ? "!" : "") + "}" : "0") + O + " + " + (R ? "{" + (C.length + (Q ? 1 : 0)) + (F ? "!" : "") + "}" : "0") + P + ")", Q && (C.push(Q), D.push(0)), R && (C.push(0), D.push(R))
                                        }
                                    } else {
                                        if (H !== I) {
                                            E = 0;
                                            break
                                        }
                                        g += H, z++, B++, 0 === E && "c" === H || 1 === E && "a" === H || 2 === E && "l" === H || 3 === E && "c" === H || E >= 4 && "(" === H ? E++ : (E && E < 5 || E >= 4 && ")" === H && --E < 5) && (E = 0), 0 === F && "r" === H || 1 === F && "g" === H || 2 === F && "b" === H || 3 === F && "a" === H || F >= 3 && "(" === H ? (3 === F && "a" === H && (G = 1), F++) : G && "," === H ? ++G > 3 && (F = G = 0) : (G && F < (G ? 5 : 4) || F >= (G ? 4 : 3) && ")" === H && --F < (G ? 5 : 4)) && (F = G = 0)
                                    }
                                }
                                z === r.length && B === p.length || (y.debug && console.error('Trying to pattern match mis-matched strings ["' + p + '", "' + r + '"]'), g = d), g && (C.length ? (y.debug && console.log('Pattern found "' + g + '" -> ', C, D, "[" + r + "," + p + "]"), r = C, p = D, t = v = "") : g = d)
                            }
                            g || (s = x(e, r), r = s[0], v = s[1], s = x(e, p), p = s[0].replace(/^([+-\/*])=/, function(a, b) {
                                return w = b, ""
                            }), t = s[1], r = parseFloat(r) || 0, p = parseFloat(p) || 0, "%" === t && (/^(fontSize|lineHeight)$/.test(e) ? (p /= 100, t = "em") : /^scale/.test(e) ? (p /= 100, t = "") : /(Red|Green|Blue)$/i.test(e) && (p = p / 100 * 255, t = "")));
                            if (/[\/*]/.test(w)) t = v;
                            else if (v !== t && 0 !== r)
                                if (0 === p) t = v;
                                else {
                                    h = h || function() {
                                        var d = {
                                                myParent: a.parentNode || c.body,
                                                position: A.getPropertyValue(a, "position"),
                                                fontSize: A.getPropertyValue(a, "fontSize")
                                            },
                                            e = d.position === M.lastPosition && d.myParent === M.lastParent,
                                            f = d.fontSize === M.lastFontSize;
                                        M.lastParent = d.myParent, M.lastPosition = d.position, M.lastFontSize = d.fontSize;
                                        var g = {};
                                        if (f && e) g.emToPx = M.lastEmToPx, g.percentToPxWidth = M.lastPercentToPxWidth, g.percentToPxHeight = M.lastPercentToPxHeight;
                                        else {
                                            var h = k && k.isSVG ? c.createElementNS("http://www.w3.org/2000/svg", "rect") : c.createElement("div");
                                            y.init(h), d.myParent.appendChild(h), o.each(["overflow", "overflowX", "overflowY"], function(a, b) {
                                                y.CSS.setPropertyValue(h, b, "hidden")
                                            }), y.CSS.setPropertyValue(h, "position", d.position), y.CSS.setPropertyValue(h, "fontSize", d.fontSize), y.CSS.setPropertyValue(h, "boxSizing", "content-box"), o.each(["minWidth", "maxWidth", "width", "minHeight", "maxHeight", "height"], function(a, b) {
                                                y.CSS.setPropertyValue(h, b, "100%")
                                            }), y.CSS.setPropertyValue(h, "paddingLeft", "100em"), g.percentToPxWidth = M.lastPercentToPxWidth = (parseFloat(A.getPropertyValue(h, "width", null, !0)) || 1) / 100, g.percentToPxHeight = M.lastPercentToPxHeight = (parseFloat(A.getPropertyValue(h, "height", null, !0)) || 1) / 100, g.emToPx = M.lastEmToPx = (parseFloat(A.getPropertyValue(h, "paddingLeft")) || 1) / 100, d.myParent.removeChild(h)
                                        }
                                        return null === M.remToPx && (M.remToPx = parseFloat(A.getPropertyValue(c.body, "fontSize")) || 16), null === M.vwToPx && (M.vwToPx = parseFloat(b.innerWidth) / 100, M.vhToPx = parseFloat(b.innerHeight) / 100), g.remToPx = M.remToPx, g.vwToPx = M.vwToPx, g.vhToPx = M.vhToPx, y.debug >= 1 && console.log("Unit ratios: " + JSON.stringify(g), a), g
                                    }();
                                    var S = /margin|padding|left|right|width|text|word|letter/i.test(e) || /X$/.test(e) || "x" === e ? "x" : "y";
                                    switch (v) {
                                        case "%":
                                            r *= "x" === S ? h.percentToPxWidth : h.percentToPxHeight;
                                            break;
                                        case "px":
                                            break;
                                        default:
                                            r *= h[v + "ToPx"]
                                    }
                                    switch (t) {
                                        case "%":
                                            r *= 1 / ("x" === S ? h.percentToPxWidth : h.percentToPxHeight);
                                            break;
                                        case "px":
                                            break;
                                        default:
                                            r *= 1 / h[t + "ToPx"]
                                    }
                                }
                            switch (w) {
                                case "+":
                                    p = r + p;
                                    break;
                                case "-":
                                    p = r - p;
                                    break;
                                case "*":
                                    p *= r;
                                    break;
                                case "/":
                                    p = r / p
                            }
                            j[e] = {
                                rootPropertyValue: m,
                                startValue: r,
                                currentValue: r,
                                endValue: p,
                                unitType: t,
                                easing: q
                            }, g && (j[e].pattern = g), y.debug && console.log("tweensContainer (" + e + "): " + JSON.stringify(j[e]), a)
                        };
                        for (var I in s)
                            if (s.hasOwnProperty(I)) {
                                var J = A.Names.camelCase(I),
                                    K = function(b, c) {
                                        var d, f, g;
                                        return u.isFunction(b) && (b = b.call(a, e, C)), u.isArray(b) ? (d = b[0], !u.isArray(b[1]) && /^[\d-]/.test(b[1]) || u.isFunction(b[1]) || A.RegEx.isHex.test(b[1]) ? g = b[1] : u.isString(b[1]) && !A.RegEx.isHex.test(b[1]) && y.Easings[b[1]] || u.isArray(b[1]) ? (f = c ? b[1] : l(b[1], i.duration), g = b[2]) : g = b[1] || b[2]) : d = b, c || (f = f || i.easing), u.isFunction(d) && (d = d.call(a, e, C)), u.isFunction(g) && (g = g.call(a, e, C)), [d || 0, f, g]
                                    }(s[I]);
                                if (t(A.Lists.colors, J)) {
                                    var L = K[0],
                                        O = K[1],
                                        P = K[2];
                                    if (A.RegEx.isHex.test(L)) {
                                        for (var Q = ["Red", "Green", "Blue"], R = A.Values.hexToRgb(L), S = P ? A.Values.hexToRgb(P) : d, T = 0; T < Q.length; T++) {
                                            var U = [R[T]];
                                            O && U.push(O), S !== d && U.push(S[T]), H(J + Q[T], U)
                                        }
                                        continue
                                    }
                                }
                                H(J, K)
                            }
                        j.element = a
                    }
                    j.element && (A.Values.addClass(a, "velocity-animating"), N.push(j), k = g(a), k && ("" === i.queue && (k.tweensContainer = j, k.opts = i), k.isAnimating = !0), D === C - 1 ? (y.State.calls.push([N, r, i, null, z.resolver, null, 0]), y.State.isTicking === !1 && (y.State.isTicking = !0, m())) : D++)
                }
                var h, i = o.extend({}, y.defaults, v),
                    j = {};
                switch (g(a) === d && y.init(a), parseFloat(i.delay) && i.queue !== !1 && o.queue(a, i.queue, function(b) {
                    y.velocityQueueEntryFlag = !0;
                    var c = y.State.delayedElements.count++;
                    y.State.delayedElements[c] = a;
                    var d = function(a) {
                        return function() {
                            y.State.delayedElements[a] = !1, b()
                        }
                    }(c);
                    g(a).delayBegin = (new Date).getTime(), g(a).delay = parseFloat(i.delay), g(a).delayTimer = {
                        setTimeout: setTimeout(b, parseFloat(i.delay)),
                        next: d
                    }
                }), i.duration.toString().toLowerCase()) {
                    case "fast":
                        i.duration = 200;
                        break;
                    case "normal":
                        i.duration = w;
                        break;
                    case "slow":
                        i.duration = 600;
                        break;
                    default:
                        i.duration = parseFloat(i.duration) || 1
                }
                if (y.mock !== !1 && (y.mock === !0 ? i.duration = i.delay = 1 : (i.duration *= parseFloat(y.mock) || 1, i.delay *= parseFloat(y.mock) || 1)), i.easing = l(i.easing, i.duration), i.begin && !u.isFunction(i.begin) && (i.begin = null), i.progress && !u.isFunction(i.progress) && (i.progress = null), i.complete && !u.isFunction(i.complete) && (i.complete = null), i.display !== d && null !== i.display && (i.display = i.display.toString().toLowerCase(), "auto" === i.display && (i.display = y.CSS.Values.getDisplayType(a))), i.visibility !== d && null !== i.visibility && (i.visibility = i.visibility.toString().toLowerCase()), i.mobileHA = i.mobileHA && y.State.isMobile && !y.State.isGingerbread, i.queue === !1)
                    if (i.delay) {
                        var k = y.State.delayedElements.count++;
                        y.State.delayedElements[k] = a;
                        var n = function(a) {
                            return function() {
                                y.State.delayedElements[a] = !1, f()
                            }
                        }(k);
                        g(a).delayBegin = (new Date).getTime(), g(a).delay = parseFloat(i.delay), g(a).delayTimer = {
                            setTimeout: setTimeout(f, parseFloat(i.delay)),
                            next: n
                        }
                    } else f();
                else o.queue(a, i.queue, function(a, b) {
                    if (b === !0) return z.promise && z.resolver(r), !0;
                    y.velocityQueueEntryFlag = !0, f(a)
                });
                "" !== i.queue && "fx" !== i.queue || "inprogress" === o.queue(a)[0] || o.dequeue(a)
            }
            var j, k, p, q, r, s, v, x = arguments[0] && (arguments[0].p || o.isPlainObject(arguments[0].properties) && !arguments[0].properties.names || u.isString(arguments[0].properties));
            u.isWrapped(this) ? (k = !1, q = 0, r = this, p = this) : (k = !0, q = 1, r = x ? arguments[0].elements || arguments[0].e : arguments[0]);
            var z = {
                promise: null,
                resolver: null,
                rejecter: null
            };
            if (k && y.Promise && (z.promise = new y.Promise(function(a, b) {
                    z.resolver = a, z.rejecter = b
                })), x ? (s = arguments[0].properties || arguments[0].p, v = arguments[0].options || arguments[0].o) : (s = arguments[q], v = arguments[q + 1]), !(r = f(r))) return void(z.promise && (s && v && v.promiseRejectEmpty === !1 ? z.resolver() : z.rejecter()));
            var C = r.length,
                D = 0;
            if (!/^(stop|finish|finishAll|pause|resume)$/i.test(s) && !o.isPlainObject(v)) {
                var E = q + 1;
                v = {};
                for (var F = E; F < arguments.length; F++) u.isArray(arguments[F]) || !/^(fast|normal|slow)$/i.test(arguments[F]) && !/^\d/.test(arguments[F]) ? u.isString(arguments[F]) || u.isArray(arguments[F]) ? v.easing = arguments[F] : u.isFunction(arguments[F]) && (v.complete = arguments[F]) : v.duration = arguments[F]
            }
            var G;
            switch (s) {
                case "scroll":
                    G = "scroll";
                    break;
                case "reverse":
                    G = "reverse";
                    break;
                case "pause":
                    var H = (new Date).getTime();
                    return o.each(r, function(a, b) {
                        h(b, H)
                    }), o.each(y.State.calls, function(a, b) {
                        var c = !1;
                        b && o.each(b[1], function(a, e) {
                            var f = v === d ? "" : v;
                            return f !== !0 && b[2].queue !== f && (v !== d || b[2].queue !== !1) || (o.each(r, function(a, d) {
                                if (d === e) return b[5] = {
                                    resume: !1
                                }, c = !0, !1
                            }), !c && void 0)
                        })
                    }), a();
                case "resume":
                    return o.each(r, function(a, b) {
                        i(b, H)
                    }), o.each(y.State.calls, function(a, b) {
                        var c = !1;
                        b && o.each(b[1], function(a, e) {
                            var f = v === d ? "" : v;
                            return f !== !0 && b[2].queue !== f && (v !== d || b[2].queue !== !1) || (!b[5] || (o.each(r, function(a, d) {
                                if (d === e) return b[5].resume = !0, c = !0, !1
                            }), !c && void 0))
                        })
                    }), a();
                case "finish":
                case "finishAll":
                case "stop":
                    o.each(r, function(a, b) {
                        g(b) && g(b).delayTimer && (clearTimeout(g(b).delayTimer.setTimeout), g(b).delayTimer.next && g(b).delayTimer.next(), delete g(b).delayTimer), "finishAll" !== s || v !== !0 && !u.isString(v) || (o.each(o.queue(b, u.isString(v) ? v : ""), function(a, b) {
                            u.isFunction(b) && b()
                        }), o.queue(b, u.isString(v) ? v : "", []))
                    });
                    var I = [];
                    return o.each(y.State.calls, function(a, b) {
                        b && o.each(b[1], function(c, e) {
                            var f = v === d ? "" : v;
                            if (f !== !0 && b[2].queue !== f && (v !== d || b[2].queue !== !1)) return !0;
                            o.each(r, function(c, d) {
                                if (d === e)
                                    if ((v === !0 || u.isString(v)) && (o.each(o.queue(d, u.isString(v) ? v : ""), function(a, b) {
                                            u.isFunction(b) && b(null, !0)
                                        }), o.queue(d, u.isString(v) ? v : "", [])), "stop" === s) {
                                        var h = g(d);
                                        h && h.tweensContainer && f !== !1 && o.each(h.tweensContainer, function(a, b) {
                                            b.endValue = b.currentValue
                                        }), I.push(a)
                                    } else "finish" !== s && "finishAll" !== s || (b[2].duration = 1)
                            })
                        })
                    }), "stop" === s && (o.each(I, function(a, b) {
                        n(b, !0)
                    }), z.promise && z.resolver(r)), a();
                default:
                    if (!o.isPlainObject(s) || u.isEmptyObject(s)) {
                        if (u.isString(s) && y.Redirects[s]) {
                            j = o.extend({}, v);
                            var J = j.duration,
                                K = j.delay || 0;
                            return j.backwards === !0 && (r = o.extend(!0, [], r).reverse()), o.each(r, function(a, b) {
                                parseFloat(j.stagger) ? j.delay = K + parseFloat(j.stagger) * a : u.isFunction(j.stagger) && (j.delay = K + j.stagger.call(b, a, C)), j.drag && (j.duration = parseFloat(J) || (/^(callout|transition)/.test(s) ? 1e3 : w), j.duration = Math.max(j.duration * (j.backwards ? 1 - a / C : (a + 1) / C), .75 * j.duration, 200)), y.Redirects[s].call(b, b, j || {}, a, C, r, z.promise ? z : d)
                            }), a()
                        }
                        var L = "Velocity: First argument (" + s + ") was not a property map, a known action, or a registered redirect. Aborting.";
                        return z.promise ? z.rejecter(new Error(L)) : b.console && console.log(L), a()
                    }
                    G = "start"
            }
            var M = {
                    lastParent: null,
                    lastPosition: null,
                    lastFontSize: null,
                    lastPercentToPxWidth: null,
                    lastPercentToPxHeight: null,
                    lastEmToPx: null,
                    remToPx: null,
                    vwToPx: null,
                    vhToPx: null
                },
                N = [];
            o.each(r, function(a, b) {
                u.isNode(b) && e(b, a)
            }), j = o.extend({}, y.defaults, v), j.loop = parseInt(j.loop, 10);
            var O = 2 * j.loop - 1;
            if (j.loop)
                for (var P = 0; P < O; P++) {
                    var Q = {
                        delay: j.delay,
                        progress: j.progress
                    };
                    P === O - 1 && (Q.display = j.display, Q.visibility = j.visibility, Q.complete = j.complete), B(r, "reverse", Q)
                }
            return a()
        };
        y = o.extend(B, y), y.animate = B;
        var C = b.requestAnimationFrame || q;
        if (!y.State.isMobile && c.hidden !== d) {
            var D = function() {
                c.hidden ? (C = function(a) {
                    return setTimeout(function() {
                        a(!0)
                    }, 16)
                }, m()) : C = b.requestAnimationFrame || q
            };
            D(), c.addEventListener("visibilitychange", D)
        }
        return a.Velocity = y, a !== b && (a.fn.velocity = B, a.fn.velocity.defaults = y.defaults), o.each(["Down", "Up"], function(a, b) {
            y.Redirects["slide" + b] = function(a, c, e, f, g, h) {
                var i = o.extend({}, c),
                    j = i.begin,
                    k = i.complete,
                    l = {},
                    m = {
                        height: "",
                        marginTop: "",
                        marginBottom: "",
                        paddingTop: "",
                        paddingBottom: ""
                    };
                i.display === d && (i.display = "Down" === b ? "inline" === y.CSS.Values.getDisplayType(a) ? "inline-block" : "block" : "none"), i.begin = function() {
                    0 === e && j && j.call(g, g);
                    for (var c in m)
                        if (m.hasOwnProperty(c)) {
                            l[c] = a.style[c];
                            var d = A.getPropertyValue(a, c);
                            m[c] = "Down" === b ? [d, 0] : [0, d]
                        }
                    l.overflow = a.style.overflow, a.style.overflow = "hidden"
                }, i.complete = function() {
                    for (var b in l) l.hasOwnProperty(b) && (a.style[b] = l[b]);
                    e === f - 1 && (k && k.call(g, g), h && h.resolver(g))
                }, y(a, m, i)
            }
        }), o.each(["In", "Out"], function(a, b) {
            y.Redirects["fade" + b] = function(a, c, e, f, g, h) {
                var i = o.extend({}, c),
                    j = i.complete,
                    k = {
                        opacity: "In" === b ? 1 : 0
                    };
                0 !== e && (i.begin = null), i.complete = e !== f - 1 ? null : function() {
                    j && j.call(g, g), h && h.resolver(g)
                }, i.display === d && (i.display = "In" === b ? "auto" : "none"), y(this, k, i)
            }
        }), y
    }(window.jQuery || window.Zepto || window, window, window ? window.document : undefined)
});
! function(a) {
    "use strict";
    "function" == typeof require && "object" == typeof exports ? module.exports = a() : "function" == typeof define && define.amd ? define(["velocity"], a) : a()
}(function() {
    "use strict";
    return function(a, b, c, d) {
        var e = a.Velocity;
        if (!e || !e.Utilities) return void(b.console && console.log("Velocity UI Pack: Velocity must be loaded first. Aborting."));
        var f = e.Utilities,
            g = e.version,
            h = {
                major: 1,
                minor: 1,
                patch: 0
            };
        if (function(a, b) {
                var c = [];
                return !(!a || !b) && (f.each([a, b], function(a, b) {
                    var d = [];
                    f.each(b, function(a, b) {
                        for (; b.toString().length < 5;) b = "0" + b;
                        d.push(b)
                    }), c.push(d.join(""))
                }), parseFloat(c[0]) > parseFloat(c[1]))
            }(h, g)) {
            var i = "Velocity UI Pack: You need to update Velocity (velocity.js) to a newer version. Visit http://github.com/julianshapiro/velocity.";
            throw alert(i), new Error(i)
        }
        e.RegisterEffect = e.RegisterUI = function(a, b) {
            function c(a, b, c, d) {
                var g, h = 0;
                f.each(a.nodeType ? [a] : a, function(a, b) {
                    d && (c += a * d), g = b.parentNode;
                    var i = ["height", "paddingTop", "paddingBottom", "marginTop", "marginBottom"];
                    "border-box" === e.CSS.getPropertyValue(b, "boxSizing").toString().toLowerCase() && (i = ["height"]), f.each(i, function(a, c) {
                        h += parseFloat(e.CSS.getPropertyValue(b, c))
                    })
                }), e.animate(g, {
                    height: ("In" === b ? "+" : "-") + "=" + h
                }, {
                    queue: !1,
                    easing: "ease-in-out",
                    duration: c * ("In" === b ? .6 : 1)
                })
            }
            return e.Redirects[a] = function(d, g, h, i, j, k, l) {
                var m = h === i - 1,
                    n = 0;
                l = l || b.loop, "function" == typeof b.defaultDuration ? b.defaultDuration = b.defaultDuration.call(j, j) : b.defaultDuration = parseFloat(b.defaultDuration);
                for (var o = 0; o < b.calls.length; o++) "number" == typeof(t = b.calls[o][1]) && (n += t);
                var p = n >= 1 ? 0 : b.calls.length ? (1 - n) / b.calls.length : 1;
                for (o = 0; o < b.calls.length; o++) {
                    var q = b.calls[o],
                        r = q[0],
                        s = 1e3,
                        t = q[1],
                        u = q[2] || {},
                        v = {};
                    if (void 0 !== g.duration ? s = g.duration : void 0 !== b.defaultDuration && (s = b.defaultDuration), v.duration = s * ("number" == typeof t ? t : p), v.queue = g.queue || "", v.easing = u.easing || "ease", v.delay = parseFloat(u.delay) || 0, v.loop = !b.loop && u.loop, v._cacheValues = u._cacheValues || !0, 0 === o) {
                        if (v.delay += parseFloat(g.delay) || 0, 0 === h && (v.begin = function() {
                                g.begin && g.begin.call(j, j);
                                var b = a.match(/(In|Out)$/);
                                b && "In" === b[0] && void 0 !== r.opacity && f.each(j.nodeType ? [j] : j, function(a, b) {
                                    e.CSS.setPropertyValue(b, "opacity", 0)
                                }), g.animateParentHeight && b && c(j, b[0], s + v.delay, g.stagger)
                            }), null !== g.display)
                            if (void 0 !== g.display && "none" !== g.display) v.display = g.display;
                            else if (/In$/.test(a)) {
                            var w = e.CSS.Values.getDisplayType(d);
                            v.display = "inline" === w ? "inline-block" : w
                        }
                        g.visibility && "hidden" !== g.visibility && (v.visibility = g.visibility)
                    }
                    if (o === b.calls.length - 1) {
                        var x = function() {
                            void 0 !== g.display && "none" !== g.display || !/Out$/.test(a) || f.each(j.nodeType ? [j] : j, function(a, b) {
                                e.CSS.setPropertyValue(b, "display", "none")
                            }), g.complete && g.complete.call(j, j), k && k.resolver(j || d)
                        };
                        v.complete = function() {
                            if (l && e.Redirects[a](d, g, h, i, j, k, l === !0 || Math.max(0, l - 1)), b.reset) {
                                for (var c in b.reset)
                                    if (b.reset.hasOwnProperty(c)) {
                                        var f = b.reset[c];
                                        void 0 !== e.CSS.Hooks.registered[c] || "string" != typeof f && "number" != typeof f || (b.reset[c] = [b.reset[c], b.reset[c]])
                                    }
                                var n = {
                                    duration: 0,
                                    queue: !1
                                };
                                m && (n.complete = x), e.animate(d, b.reset, n)
                            } else m && x()
                        }, "hidden" === g.visibility && (v.visibility = g.visibility)
                    }
                    e.animate(d, r, v)
                }
            }, e
        }, e.RegisterEffect.packagedEffects = {
            "callout.bounce": {
                defaultDuration: 550,
                calls: [
                    [{
                        translateY: -30
                    }, .25],
                    [{
                        translateY: 0
                    }, .125],
                    [{
                        translateY: -15
                    }, .125],
                    [{
                        translateY: 0
                    }, .25]
                ]
            },
            "callout.shake": {
                defaultDuration: 800,
                calls: [
                    [{
                        translateX: -11
                    }],
                    [{
                        translateX: 11
                    }],
                    [{
                        translateX: -11
                    }],
                    [{
                        translateX: 11
                    }],
                    [{
                        translateX: -11
                    }],
                    [{
                        translateX: 11
                    }],
                    [{
                        translateX: -11
                    }],
                    [{
                        translateX: 0
                    }]
                ]
            },
            "callout.flash": {
                defaultDuration: 1100,
                calls: [
                    [{
                        opacity: [0, "easeInOutQuad", 1]
                    }],
                    [{
                        opacity: [1, "easeInOutQuad"]
                    }],
                    [{
                        opacity: [0, "easeInOutQuad"]
                    }],
                    [{
                        opacity: [1, "easeInOutQuad"]
                    }]
                ]
            },
            "callout.pulse": {
                defaultDuration: 825,
                calls: [
                    [{
                        scaleX: 1.1,
                        scaleY: 1.1
                    }, .5, {
                        easing: "easeInExpo"
                    }],
                    [{
                        scaleX: 1,
                        scaleY: 1
                    }, .5]
                ]
            },
            "callout.swing": {
                defaultDuration: 950,
                calls: [
                    [{
                        rotateZ: 15
                    }],
                    [{
                        rotateZ: -10
                    }],
                    [{
                        rotateZ: 5
                    }],
                    [{
                        rotateZ: -5
                    }],
                    [{
                        rotateZ: 0
                    }]
                ]
            },
            "callout.tada": {
                defaultDuration: 1e3,
                calls: [
                    [{
                        scaleX: .9,
                        scaleY: .9,
                        rotateZ: -3
                    }, .1],
                    [{
                        scaleX: 1.1,
                        scaleY: 1.1,
                        rotateZ: 3
                    }, .1],
                    [{
                        scaleX: 1.1,
                        scaleY: 1.1,
                        rotateZ: -3
                    }, .1],
                    ["reverse", .125],
                    ["reverse", .125],
                    ["reverse", .125],
                    ["reverse", .125],
                    ["reverse", .125],
                    [{
                        scaleX: 1,
                        scaleY: 1,
                        rotateZ: 0
                    }, .2]
                ]
            },
            "transition.fadeIn": {
                defaultDuration: 500,
                calls: [
                    [{
                        opacity: [1, 0]
                    }]
                ]
            },
            "transition.fadeOut": {
                defaultDuration: 500,
                calls: [
                    [{
                        opacity: [0, 1]
                    }]
                ]
            },
            "transition.flipXIn": {
                defaultDuration: 700,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformPerspective: [800, 800],
                        rotateY: [0, -55]
                    }]
                ],
                reset: {
                    transformPerspective: 0
                }
            },
            "transition.flipXOut": {
                defaultDuration: 700,
                calls: [
                    [{
                        opacity: [0, 1],
                        transformPerspective: [800, 800],
                        rotateY: 55
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    rotateY: 0
                }
            },
            "transition.flipYIn": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformPerspective: [800, 800],
                        rotateX: [0, -45]
                    }]
                ],
                reset: {
                    transformPerspective: 0
                }
            },
            "transition.flipYOut": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [0, 1],
                        transformPerspective: [800, 800],
                        rotateX: 25
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    rotateX: 0
                }
            },
            "transition.flipBounceXIn": {
                defaultDuration: 900,
                calls: [
                    [{
                        opacity: [.725, 0],
                        transformPerspective: [400, 400],
                        rotateY: [-10, 90]
                    }, .5],
                    [{
                        opacity: .8,
                        rotateY: 10
                    }, .25],
                    [{
                        opacity: 1,
                        rotateY: 0
                    }, .25]
                ],
                reset: {
                    transformPerspective: 0
                }
            },
            "transition.flipBounceXOut": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [.9, 1],
                        transformPerspective: [400, 400],
                        rotateY: -10
                    }],
                    [{
                        opacity: 0,
                        rotateY: 90
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    rotateY: 0
                }
            },
            "transition.flipBounceYIn": {
                defaultDuration: 850,
                calls: [
                    [{
                        opacity: [.725, 0],
                        transformPerspective: [400, 400],
                        rotateX: [-10, 90]
                    }, .5],
                    [{
                        opacity: .8,
                        rotateX: 10
                    }, .25],
                    [{
                        opacity: 1,
                        rotateX: 0
                    }, .25]
                ],
                reset: {
                    transformPerspective: 0
                }
            },
            "transition.flipBounceYOut": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [.9, 1],
                        transformPerspective: [400, 400],
                        rotateX: -15
                    }],
                    [{
                        opacity: 0,
                        rotateX: 90
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    rotateX: 0
                }
            },
            "transition.swoopIn": {
                defaultDuration: 850,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformOriginX: ["100%", "50%"],
                        transformOriginY: ["100%", "100%"],
                        scaleX: [1, 0],
                        scaleY: [1, 0],
                        translateX: [0, -700],
                        translateZ: 0
                    }]
                ],
                reset: {
                    transformOriginX: "50%",
                    transformOriginY: "50%"
                }
            },
            "transition.swoopOut": {
                defaultDuration: 850,
                calls: [
                    [{
                        opacity: [0, 1],
                        transformOriginX: ["50%", "100%"],
                        transformOriginY: ["100%", "100%"],
                        scaleX: 0,
                        scaleY: 0,
                        translateX: -700,
                        translateZ: 0
                    }]
                ],
                reset: {
                    transformOriginX: "50%",
                    transformOriginY: "50%",
                    scaleX: 1,
                    scaleY: 1,
                    translateX: 0
                }
            },
            "transition.whirlIn": {
                defaultDuration: 850,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformOriginX: ["50%", "50%"],
                        transformOriginY: ["50%", "50%"],
                        scaleX: [1, 0],
                        scaleY: [1, 0],
                        rotateY: [0, 160]
                    }, 1, {
                        easing: "easeInOutSine"
                    }]
                ]
            },
            "transition.whirlOut": {
                defaultDuration: 750,
                calls: [
                    [{
                        opacity: [0, "easeInOutQuint", 1],
                        transformOriginX: ["50%", "50%"],
                        transformOriginY: ["50%", "50%"],
                        scaleX: 0,
                        scaleY: 0,
                        rotateY: 160
                    }, 1, {
                        easing: "swing"
                    }]
                ],
                reset: {
                    scaleX: 1,
                    scaleY: 1,
                    rotateY: 0
                }
            },
            "transition.shrinkIn": {
                defaultDuration: 750,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformOriginX: ["50%", "50%"],
                        transformOriginY: ["50%", "50%"],
                        scaleX: [1, 1.5],
                        scaleY: [1, 1.5],
                        translateZ: 0
                    }]
                ]
            },
            "transition.shrinkOut": {
                defaultDuration: 600,
                calls: [
                    [{
                        opacity: [0, 1],
                        transformOriginX: ["50%", "50%"],
                        transformOriginY: ["50%", "50%"],
                        scaleX: 1.3,
                        scaleY: 1.3,
                        translateZ: 0
                    }]
                ],
                reset: {
                    scaleX: 1,
                    scaleY: 1
                }
            },
            "transition.expandIn": {
                defaultDuration: 700,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformOriginX: ["50%", "50%"],
                        transformOriginY: ["50%", "50%"],
                        scaleX: [1, .625],
                        scaleY: [1, .625],
                        translateZ: 0
                    }]
                ]
            },
            "transition.expandOut": {
                defaultDuration: 700,
                calls: [
                    [{
                        opacity: [0, 1],
                        transformOriginX: ["50%", "50%"],
                        transformOriginY: ["50%", "50%"],
                        scaleX: .5,
                        scaleY: .5,
                        translateZ: 0
                    }]
                ],
                reset: {
                    scaleX: 1,
                    scaleY: 1
                }
            },
            "transition.bounceIn": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [1, 0],
                        scaleX: [1.05, .3],
                        scaleY: [1.05, .3]
                    }, .35],
                    [{
                        scaleX: .9,
                        scaleY: .9,
                        translateZ: 0
                    }, .2],
                    [{
                        scaleX: 1,
                        scaleY: 1
                    }, .45]
                ]
            },
            "transition.bounceOut": {
                defaultDuration: 800,
                calls: [
                    [{
                        scaleX: .95,
                        scaleY: .95
                    }, .35],
                    [{
                        scaleX: 1.1,
                        scaleY: 1.1,
                        translateZ: 0
                    }, .35],
                    [{
                        opacity: [0, 1],
                        scaleX: .3,
                        scaleY: .3
                    }, .3]
                ],
                reset: {
                    scaleX: 1,
                    scaleY: 1
                }
            },
            "transition.bounceUpIn": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateY: [-30, 1e3]
                    }, .6, {
                        easing: "easeOutCirc"
                    }],
                    [{
                        translateY: 10
                    }, .2],
                    [{
                        translateY: 0
                    }, .2]
                ]
            },
            "transition.bounceUpOut": {
                defaultDuration: 1e3,
                calls: [
                    [{
                        translateY: 20
                    }, .2],
                    [{
                        opacity: [0, "easeInCirc", 1],
                        translateY: -1e3
                    }, .8]
                ],
                reset: {
                    translateY: 0
                }
            },
            "transition.bounceDownIn": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateY: [30, -1e3]
                    }, .6, {
                        easing: "easeOutCirc"
                    }],
                    [{
                        translateY: -10
                    }, .2],
                    [{
                        translateY: 0
                    }, .2]
                ]
            },
            "transition.bounceDownOut": {
                defaultDuration: 1e3,
                calls: [
                    [{
                        translateY: -20
                    }, .2],
                    [{
                        opacity: [0, "easeInCirc", 1],
                        translateY: 1e3
                    }, .8]
                ],
                reset: {
                    translateY: 0
                }
            },
            "transition.bounceLeftIn": {
                defaultDuration: 750,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateX: [30, -1250]
                    }, .6, {
                        easing: "easeOutCirc"
                    }],
                    [{
                        translateX: -10
                    }, .2],
                    [{
                        translateX: 0
                    }, .2]
                ]
            },
            "transition.bounceLeftOut": {
                defaultDuration: 750,
                calls: [
                    [{
                        translateX: 30
                    }, .2],
                    [{
                        opacity: [0, "easeInCirc", 1],
                        translateX: -1250
                    }, .8]
                ],
                reset: {
                    translateX: 0
                }
            },
            "transition.bounceRightIn": {
                defaultDuration: 750,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateX: [-30, 1250]
                    }, .6, {
                        easing: "easeOutCirc"
                    }],
                    [{
                        translateX: 10
                    }, .2],
                    [{
                        translateX: 0
                    }, .2]
                ]
            },
            "transition.bounceRightOut": {
                defaultDuration: 750,
                calls: [
                    [{
                        translateX: -30
                    }, .2],
                    [{
                        opacity: [0, "easeInCirc", 1],
                        translateX: 1250
                    }, .8]
                ],
                reset: {
                    translateX: 0
                }
            },
            "transition.slideUpIn": {
                defaultDuration: 900,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateY: [0, 20],
                        translateZ: 0
                    }]
                ]
            },
            "transition.slideUpOut": {
                defaultDuration: 900,
                calls: [
                    [{
                        opacity: [0, 1],
                        translateY: -20,
                        translateZ: 0
                    }]
                ],
                reset: {
                    translateY: 0
                }
            },
            "transition.slideDownIn": {
                defaultDuration: 900,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateY: [0, -20],
                        translateZ: 0
                    }]
                ]
            },
            "transition.slideDownOut": {
                defaultDuration: 900,
                calls: [
                    [{
                        opacity: [0, 1],
                        translateY: 20,
                        translateZ: 0
                    }]
                ],
                reset: {
                    translateY: 0
                }
            },
            "transition.slideLeftIn": {
                defaultDuration: 1e3,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateX: [0, -20],
                        translateZ: 0
                    }]
                ]
            },
            "transition.slideLeftOut": {
                defaultDuration: 1050,
                calls: [
                    [{
                        opacity: [0, 1],
                        translateX: -20,
                        translateZ: 0
                    }]
                ],
                reset: {
                    translateX: 0
                }
            },
            "transition.slideRightIn": {
                defaultDuration: 1e3,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateX: [0, 20],
                        translateZ: 0
                    }]
                ]
            },
            "transition.slideRightOut": {
                defaultDuration: 1050,
                calls: [
                    [{
                        opacity: [0, 1],
                        translateX: 20,
                        translateZ: 0
                    }]
                ],
                reset: {
                    translateX: 0
                }
            },
            "transition.slideUpBigIn": {
                defaultDuration: 850,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateY: [0, 75],
                        translateZ: 0
                    }]
                ]
            },
            "transition.slideUpBigOut": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [0, 1],
                        translateY: -75,
                        translateZ: 0
                    }]
                ],
                reset: {
                    translateY: 0
                }
            },
            "transition.slideDownBigIn": {
                defaultDuration: 850,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateY: [0, -75],
                        translateZ: 0
                    }]
                ]
            },
            "transition.slideDownBigOut": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [0, 1],
                        translateY: 75,
                        translateZ: 0
                    }]
                ],
                reset: {
                    translateY: 0
                }
            },
            "transition.slideLeftBigIn": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateX: [0, -75],
                        translateZ: 0
                    }]
                ]
            },
            "transition.slideLeftBigOut": {
                defaultDuration: 750,
                calls: [
                    [{
                        opacity: [0, 1],
                        translateX: -75,
                        translateZ: 0
                    }]
                ],
                reset: {
                    translateX: 0
                }
            },
            "transition.slideRightBigIn": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [1, 0],
                        translateX: [0, 75],
                        translateZ: 0
                    }]
                ]
            },
            "transition.slideRightBigOut": {
                defaultDuration: 750,
                calls: [
                    [{
                        opacity: [0, 1],
                        translateX: 75,
                        translateZ: 0
                    }]
                ],
                reset: {
                    translateX: 0
                }
            },
            "transition.perspectiveUpIn": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformPerspective: [800, 800],
                        transformOriginX: [0, 0],
                        transformOriginY: ["100%", "100%"],
                        rotateX: [0, -180]
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    transformOriginX: "50%",
                    transformOriginY: "50%"
                }
            },
            "transition.perspectiveUpOut": {
                defaultDuration: 850,
                calls: [
                    [{
                        opacity: [0, 1],
                        transformPerspective: [800, 800],
                        transformOriginX: [0, 0],
                        transformOriginY: ["100%", "100%"],
                        rotateX: -180
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    transformOriginX: "50%",
                    transformOriginY: "50%",
                    rotateX: 0
                }
            },
            "transition.perspectiveDownIn": {
                defaultDuration: 800,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformPerspective: [800, 800],
                        transformOriginX: [0, 0],
                        transformOriginY: [0, 0],
                        rotateX: [0, 180]
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    transformOriginX: "50%",
                    transformOriginY: "50%"
                }
            },
            "transition.perspectiveDownOut": {
                defaultDuration: 850,
                calls: [
                    [{
                        opacity: [0, 1],
                        transformPerspective: [800, 800],
                        transformOriginX: [0, 0],
                        transformOriginY: [0, 0],
                        rotateX: 180
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    transformOriginX: "50%",
                    transformOriginY: "50%",
                    rotateX: 0
                }
            },
            "transition.perspectiveLeftIn": {
                defaultDuration: 950,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformPerspective: [2e3, 2e3],
                        transformOriginX: [0, 0],
                        transformOriginY: [0, 0],
                        rotateY: [0, -180]
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    transformOriginX: "50%",
                    transformOriginY: "50%"
                }
            },
            "transition.perspectiveLeftOut": {
                defaultDuration: 950,
                calls: [
                    [{
                        opacity: [0, 1],
                        transformPerspective: [2e3, 2e3],
                        transformOriginX: [0, 0],
                        transformOriginY: [0, 0],
                        rotateY: -180
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    transformOriginX: "50%",
                    transformOriginY: "50%",
                    rotateY: 0
                }
            },
            "transition.perspectiveRightIn": {
                defaultDuration: 950,
                calls: [
                    [{
                        opacity: [1, 0],
                        transformPerspective: [2e3, 2e3],
                        transformOriginX: ["100%", "100%"],
                        transformOriginY: [0, 0],
                        rotateY: [0, 180]
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    transformOriginX: "50%",
                    transformOriginY: "50%"
                }
            },
            "transition.perspectiveRightOut": {
                defaultDuration: 950,
                calls: [
                    [{
                        opacity: [0, 1],
                        transformPerspective: [2e3, 2e3],
                        transformOriginX: ["100%", "100%"],
                        transformOriginY: [0, 0],
                        rotateY: 180
                    }]
                ],
                reset: {
                    transformPerspective: 0,
                    transformOriginX: "50%",
                    transformOriginY: "50%",
                    rotateY: 0
                }
            }
        };
        for (var j in e.RegisterEffect.packagedEffects) e.RegisterEffect.packagedEffects.hasOwnProperty(j) && e.RegisterEffect(j, e.RegisterEffect.packagedEffects[j]);
        e.RunSequence = function(a) {
            var b = f.extend(!0, [], a);
            b.length > 1 && (f.each(b.reverse(), function(a, c) {
                var d = b[a + 1];
                if (d) {
                    var g = c.o || c.options,
                        h = d.o || d.options,
                        i = g && g.sequenceQueue === !1 ? "begin" : "complete",
                        j = h && h[i],
                        k = {};
                    k[i] = function() {
                        var a = d.e || d.elements,
                            b = a.nodeType ? [a] : a;
                        j && j.call(b, b), e(c)
                    }, d.o ? d.o = f.extend({}, h, k) : d.options = f.extend({}, h, k)
                }
            }), b.reverse()), e(b[0])
        }
    }(window.jQuery || window.Zepto || window, window, window ? window.document : undefined)
});
! function(i) {
    "use strict";
    "function" == typeof define && define.amd ? define(["jquery"], i) : "undefined" != typeof exports ? module.exports = i(require("jquery")) : i(jQuery)
}(function(i) {
    "use strict";
    var e = window.Slick || {};
    (e = function() {
        var e = 0;
        return function(t, o) {
            var s, n = this;
            n.defaults = {
                accessibility: !0,
                adaptiveHeight: !1,
                appendArrows: i(t),
                appendDots: i(t),
                arrows: !0,
                asNavFor: null,
                prevArrow: '<button class="slick-prev" aria-label="Previous" type="button">Previous</button>',
                nextArrow: '<button class="slick-next" aria-label="Next" type="button">Next</button>',
                autoplay: !1,
                autoplaySpeed: 3e3,
                centerMode: !1,
                centerPadding: "50px",
                cssEase: "ease",
                customPaging: function(e, t) {
                    return i('<button type="button" />').text(t + 1)
                },
                dots: !1,
                dotsClass: "slick-dots",
                draggable: !0,
                easing: "linear",
                edgeFriction: .35,
                fade: !1,
                focusOnSelect: !1,
                focusOnChange: !1,
                infinite: !0,
                initialSlide: 0,
                lazyLoad: "ondemand",
                mobileFirst: !1,
                pauseOnHover: !0,
                pauseOnFocus: !0,
                pauseOnDotsHover: !1,
                respondTo: "window",
                responsive: null,
                rows: 1,
                rtl: !1,
                slide: "",
                slidesPerRow: 1,
                slidesToShow: 1,
                slidesToScroll: 1,
                speed: 500,
                swipe: !0,
                swipeToSlide: !1,
                touchMove: !0,
                touchThreshold: 5,
                useCSS: !0,
                useTransform: !0,
                variableWidth: !1,
                vertical: !1,
                verticalSwiping: !1,
                waitForAnimate: !0,
                zIndex: 1e3
            }, n.initials = {
                animating: !1,
                dragging: !1,
                autoPlayTimer: null,
                currentDirection: 0,
                currentLeft: null,
                currentSlide: 0,
                direction: 1,
                $dots: null,
                listWidth: null,
                listHeight: null,
                loadIndex: 0,
                $nextArrow: null,
                $prevArrow: null,
                scrolling: !1,
                slideCount: null,
                slideWidth: null,
                $slideTrack: null,
                $slides: null,
                sliding: !1,
                slideOffset: 0,
                swipeLeft: null,
                swiping: !1,
                $list: null,
                touchObject: {},
                transformsEnabled: !1,
                unslicked: !1
            }, i.extend(n, n.initials), n.activeBreakpoint = null, n.animType = null, n.animProp = null, n.breakpoints = [], n.breakpointSettings = [], n.cssTransitions = !1, n.focussed = !1, n.interrupted = !1, n.hidden = "hidden", n.paused = !0, n.positionProp = null, n.respondTo = null, n.rowCount = 1, n.shouldClick = !0, n.$slider = i(t), n.$slidesCache = null, n.transformType = null, n.transitionType = null, n.visibilityChange = "visibilitychange", n.windowWidth = 0, n.windowTimer = null, s = i(t).data("slick") || {}, n.options = i.extend({}, n.defaults, o, s), n.currentSlide = n.options.initialSlide, n.originalSettings = n.options, void 0 !== document.mozHidden ? (n.hidden = "mozHidden", n.visibilityChange = "mozvisibilitychange") : void 0 !== document.webkitHidden && (n.hidden = "webkitHidden", n.visibilityChange = "webkitvisibilitychange"), n.autoPlay = i.proxy(n.autoPlay, n), n.autoPlayClear = i.proxy(n.autoPlayClear, n), n.autoPlayIterator = i.proxy(n.autoPlayIterator, n), n.changeSlide = i.proxy(n.changeSlide, n), n.clickHandler = i.proxy(n.clickHandler, n), n.selectHandler = i.proxy(n.selectHandler, n), n.setPosition = i.proxy(n.setPosition, n), n.swipeHandler = i.proxy(n.swipeHandler, n), n.dragHandler = i.proxy(n.dragHandler, n), n.keyHandler = i.proxy(n.keyHandler, n), n.instanceUid = e++, n.htmlExpr = /^(?:\s*(<[\w\W]+>)[^>]*)$/, n.registerBreakpoints(), n.init(!0)
        }
    }()).prototype.activateADA = function() {
        this.$slideTrack.find(".slick-active").attr({
            "aria-hidden": "false"
        }).find("a, input, button, select").attr({
            tabindex: "0"
        })
    }, e.prototype.addSlide = e.prototype.slickAdd = function(e, t, o) {
        var s = this;
        if ("boolean" == typeof t) o = t, t = null;
        else if (t < 0 || t >= s.slideCount) return !1;
        s.unload(), "number" == typeof t ? 0 === t && 0 === s.$slides.length ? i(e).appendTo(s.$slideTrack) : o ? i(e).insertBefore(s.$slides.eq(t)) : i(e).insertAfter(s.$slides.eq(t)) : !0 === o ? i(e).prependTo(s.$slideTrack) : i(e).appendTo(s.$slideTrack), s.$slides = s.$slideTrack.children(this.options.slide), s.$slideTrack.children(this.options.slide).detach(), s.$slideTrack.append(s.$slides), s.$slides.each(function(e, t) {
            i(t).attr("data-slick-index", e)
        }), s.$slidesCache = s.$slides, s.reinit()
    }, e.prototype.animateHeight = function() {
        var i = this;
        if (1 === i.options.slidesToShow && !0 === i.options.adaptiveHeight && !1 === i.options.vertical) {
            var e = i.$slides.eq(i.currentSlide).outerHeight(!0);
            i.$list.animate({
                height: e
            }, i.options.speed)
        }
    }, e.prototype.animateSlide = function(e, t) {
        var o = {},
            s = this;
        s.animateHeight(), !0 === s.options.rtl && !1 === s.options.vertical && (e = -e), !1 === s.transformsEnabled ? !1 === s.options.vertical ? s.$slideTrack.animate({
            left: e
        }, s.options.speed, s.options.easing, t) : s.$slideTrack.animate({
            top: e
        }, s.options.speed, s.options.easing, t) : !1 === s.cssTransitions ? (!0 === s.options.rtl && (s.currentLeft = -s.currentLeft), i({
            animStart: s.currentLeft
        }).animate({
            animStart: e
        }, {
            duration: s.options.speed,
            easing: s.options.easing,
            step: function(i) {
                i = Math.ceil(i), !1 === s.options.vertical ? (o[s.animType] = "translate(" + i + "px, 0px)", s.$slideTrack.css(o)) : (o[s.animType] = "translate(0px," + i + "px)", s.$slideTrack.css(o))
            },
            complete: function() {
                t && t.call()
            }
        })) : (s.applyTransition(), e = Math.ceil(e), !1 === s.options.vertical ? o[s.animType] = "translate3d(" + e + "px, 0px, 0px)" : o[s.animType] = "translate3d(0px," + e + "px, 0px)", s.$slideTrack.css(o), t && setTimeout(function() {
            s.disableTransition(), t.call()
        }, s.options.speed))
    }, e.prototype.getNavTarget = function() {
        var e = this,
            t = e.options.asNavFor;
        return t && null !== t && (t = i(t).not(e.$slider)), t
    }, e.prototype.asNavFor = function(e) {
        var t = this.getNavTarget();
        null !== t && "object" == typeof t && t.each(function() {
            var t = i(this).slick("getSlick");
            t.unslicked || t.slideHandler(e, !0)
        })
    }, e.prototype.applyTransition = function(i) {
        var e = this,
            t = {};
        !1 === e.options.fade ? t[e.transitionType] = e.transformType + " " + e.options.speed + "ms " + e.options.cssEase : t[e.transitionType] = "opacity " + e.options.speed + "ms " + e.options.cssEase, !1 === e.options.fade ? e.$slideTrack.css(t) : e.$slides.eq(i).css(t)
    }, e.prototype.autoPlay = function() {
        var i = this;
        i.autoPlayClear(), i.slideCount > i.options.slidesToShow && (i.autoPlayTimer = setInterval(i.autoPlayIterator, i.options.autoplaySpeed))
    }, e.prototype.autoPlayClear = function() {
        var i = this;
        i.autoPlayTimer && clearInterval(i.autoPlayTimer)
    }, e.prototype.autoPlayIterator = function() {
        var i = this,
            e = i.currentSlide + i.options.slidesToScroll;
        i.paused || i.interrupted || i.focussed || (!1 === i.options.infinite && (1 === i.direction && i.currentSlide + 1 === i.slideCount - 1 ? i.direction = 0 : 0 === i.direction && (e = i.currentSlide - i.options.slidesToScroll, i.currentSlide - 1 == 0 && (i.direction = 1))), i.slideHandler(e))
    }, e.prototype.buildArrows = function() {
        var e = this;
        !0 === e.options.arrows && (e.$prevArrow = i(e.options.prevArrow).addClass("slick-arrow"), e.$nextArrow = i(e.options.nextArrow).addClass("slick-arrow"), e.slideCount > e.options.slidesToShow ? (e.$prevArrow.removeClass("slick-hidden").removeAttr("aria-hidden tabindex"), e.$nextArrow.removeClass("slick-hidden").removeAttr("aria-hidden tabindex"), e.htmlExpr.test(e.options.prevArrow) && e.$prevArrow.prependTo(e.options.appendArrows), e.htmlExpr.test(e.options.nextArrow) && e.$nextArrow.appendTo(e.options.appendArrows), !0 !== e.options.infinite && e.$prevArrow.addClass("slick-disabled").attr("aria-disabled", "true")) : e.$prevArrow.add(e.$nextArrow).addClass("slick-hidden").attr({
            "aria-disabled": "true",
            tabindex: "-1"
        }))
    }, e.prototype.buildDots = function() {
        var e, t, o = this;
        if (!0 === o.options.dots) {
            for (o.$slider.addClass("slick-dotted"), t = i("<ul />").addClass(o.options.dotsClass), e = 0; e <= o.getDotCount(); e += 1) t.append(i("<li />").append(o.options.customPaging.call(this, o, e)));
            o.$dots = t.appendTo(o.options.appendDots), o.$dots.find("li").first().addClass("slick-active")
        }
    }, e.prototype.buildOut = function() {
        var e = this;
        e.$slides = e.$slider.children(e.options.slide + ":not(.slick-cloned)").addClass("slick-slide"), e.slideCount = e.$slides.length, e.$slides.each(function(e, t) {
            i(t).attr("data-slick-index", e).data("originalStyling", i(t).attr("style") || "")
        }), e.$slider.addClass("slick-slider"), e.$slideTrack = 0 === e.slideCount ? i('<div class="slick-track"/>').appendTo(e.$slider) : e.$slides.wrapAll('<div class="slick-track"/>').parent(), e.$list = e.$slideTrack.wrap('<div class="slick-list"/>').parent(), e.$slideTrack.css("opacity", 0), !0 !== e.options.centerMode && !0 !== e.options.swipeToSlide || (e.options.slidesToScroll = 1), i("img[data-lazy]", e.$slider).not("[src]").addClass("slick-loading"), e.setupInfinite(), e.buildArrows(), e.buildDots(), e.updateDots(), e.setSlideClasses("number" == typeof e.currentSlide ? e.currentSlide : 0), !0 === e.options.draggable && e.$list.addClass("draggable")
    }, e.prototype.buildRows = function() {
        var i, e, t, o, s, n, r, l = this;
        if (o = document.createDocumentFragment(), n = l.$slider.children(), l.options.rows > 1) {
            for (r = l.options.slidesPerRow * l.options.rows, s = Math.ceil(n.length / r), i = 0; i < s; i++) {
                var d = document.createElement("div");
                for (e = 0; e < l.options.rows; e++) {
                    var a = document.createElement("div");
                    for (t = 0; t < l.options.slidesPerRow; t++) {
                        var c = i * r + (e * l.options.slidesPerRow + t);
                        n.get(c) && a.appendChild(n.get(c))
                    }
                    d.appendChild(a)
                }
                o.appendChild(d)
            }
            l.$slider.empty().append(o), l.$slider.children().children().children().css({
                width: 100 / l.options.slidesPerRow + "%",
                display: "inline-block"
            })
        }
    }, e.prototype.checkResponsive = function(e, t) {
        var o, s, n, r = this,
            l = !1,
            d = r.$slider.width(),
            a = window.innerWidth || i(window).width();
        if ("window" === r.respondTo ? n = a : "slider" === r.respondTo ? n = d : "min" === r.respondTo && (n = Math.min(a, d)), r.options.responsive && r.options.responsive.length && null !== r.options.responsive) {
            s = null;
            for (o in r.breakpoints) r.breakpoints.hasOwnProperty(o) && (!1 === r.originalSettings.mobileFirst ? n < r.breakpoints[o] && (s = r.breakpoints[o]) : n > r.breakpoints[o] && (s = r.breakpoints[o]));
            null !== s ? null !== r.activeBreakpoint ? (s !== r.activeBreakpoint || t) && (r.activeBreakpoint = s, "unslick" === r.breakpointSettings[s] ? r.unslick(s) : (r.options = i.extend({}, r.originalSettings, r.breakpointSettings[s]), !0 === e && (r.currentSlide = r.options.initialSlide), r.refresh(e)), l = s) : (r.activeBreakpoint = s, "unslick" === r.breakpointSettings[s] ? r.unslick(s) : (r.options = i.extend({}, r.originalSettings, r.breakpointSettings[s]), !0 === e && (r.currentSlide = r.options.initialSlide), r.refresh(e)), l = s) : null !== r.activeBreakpoint && (r.activeBreakpoint = null, r.options = r.originalSettings, !0 === e && (r.currentSlide = r.options.initialSlide), r.refresh(e), l = s), e || !1 === l || r.$slider.trigger("breakpoint", [r, l])
        }
    }, e.prototype.changeSlide = function(e, t) {
        var o, s, n, r = this,
            l = i(e.currentTarget);
        switch (l.is("a") && e.preventDefault(), l.is("li") || (l = l.closest("li")), n = r.slideCount % r.options.slidesToScroll != 0, o = n ? 0 : (r.slideCount - r.currentSlide) % r.options.slidesToScroll, e.data.message) {
            case "previous":
                s = 0 === o ? r.options.slidesToScroll : r.options.slidesToShow - o, r.slideCount > r.options.slidesToShow && r.slideHandler(r.currentSlide - s, !1, t);
                break;
            case "next":
                s = 0 === o ? r.options.slidesToScroll : o, r.slideCount > r.options.slidesToShow && r.slideHandler(r.currentSlide + s, !1, t);
                break;
            case "index":
                var d = 0 === e.data.index ? 0 : e.data.index || l.index() * r.options.slidesToScroll;
                r.slideHandler(r.checkNavigable(d), !1, t), l.children().trigger("focus");
                break;
            default:
                return
        }
    }, e.prototype.checkNavigable = function(i) {
        var e, t;
        if (e = this.getNavigableIndexes(), t = 0, i > e[e.length - 1]) i = e[e.length - 1];
        else
            for (var o in e) {
                if (i < e[o]) {
                    i = t;
                    break
                }
                t = e[o]
            }
        return i
    }, e.prototype.cleanUpEvents = function() {
        var e = this;
        e.options.dots && null !== e.$dots && (i("li", e.$dots).off("click.slick", e.changeSlide).off("mouseenter.slick", i.proxy(e.interrupt, e, !0)).off("mouseleave.slick", i.proxy(e.interrupt, e, !1)), !0 === e.options.accessibility && e.$dots.off("keydown.slick", e.keyHandler)), e.$slider.off("focus.slick blur.slick"), !0 === e.options.arrows && e.slideCount > e.options.slidesToShow && (e.$prevArrow && e.$prevArrow.off("click.slick", e.changeSlide), e.$nextArrow && e.$nextArrow.off("click.slick", e.changeSlide), !0 === e.options.accessibility && (e.$prevArrow && e.$prevArrow.off("keydown.slick", e.keyHandler), e.$nextArrow && e.$nextArrow.off("keydown.slick", e.keyHandler))), e.$list.off("touchstart.slick mousedown.slick", e.swipeHandler), e.$list.off("touchmove.slick mousemove.slick", e.swipeHandler), e.$list.off("touchend.slick mouseup.slick", e.swipeHandler), e.$list.off("touchcancel.slick mouseleave.slick", e.swipeHandler), e.$list.off("click.slick", e.clickHandler), i(document).off(e.visibilityChange, e.visibility), e.cleanUpSlideEvents(), !0 === e.options.accessibility && e.$list.off("keydown.slick", e.keyHandler), !0 === e.options.focusOnSelect && i(e.$slideTrack).children().off("click.slick", e.selectHandler), i(window).off("orientationchange.slick.slick-" + e.instanceUid, e.orientationChange), i(window).off("resize.slick.slick-" + e.instanceUid, e.resize), i("[draggable!=true]", e.$slideTrack).off("dragstart", e.preventDefault), i(window).off("load.slick.slick-" + e.instanceUid, e.setPosition)
    }, e.prototype.cleanUpSlideEvents = function() {
        var e = this;
        e.$list.off("mouseenter.slick", i.proxy(e.interrupt, e, !0)), e.$list.off("mouseleave.slick", i.proxy(e.interrupt, e, !1))
    }, e.prototype.cleanUpRows = function() {
        var i, e = this;
        e.options.rows > 1 && ((i = e.$slides.children().children()).removeAttr("style"), e.$slider.empty().append(i))
    }, e.prototype.clickHandler = function(i) {
        !1 === this.shouldClick && (i.stopImmediatePropagation(), i.stopPropagation(), i.preventDefault())
    }, e.prototype.destroy = function(e) {
        var t = this;
        t.autoPlayClear(), t.touchObject = {}, t.cleanUpEvents(), i(".slick-cloned", t.$slider).detach(), t.$dots && t.$dots.remove(), t.$prevArrow && t.$prevArrow.length && (t.$prevArrow.removeClass("slick-disabled slick-arrow slick-hidden").removeAttr("aria-hidden aria-disabled tabindex").css("display", ""), t.htmlExpr.test(t.options.prevArrow) && t.$prevArrow.remove()), t.$nextArrow && t.$nextArrow.length && (t.$nextArrow.removeClass("slick-disabled slick-arrow slick-hidden").removeAttr("aria-hidden aria-disabled tabindex").css("display", ""), t.htmlExpr.test(t.options.nextArrow) && t.$nextArrow.remove()), t.$slides && (t.$slides.removeClass("slick-slide slick-active slick-center slick-visible slick-current").removeAttr("aria-hidden").removeAttr("data-slick-index").each(function() {
            i(this).attr("style", i(this).data("originalStyling"))
        }), t.$slideTrack.children(this.options.slide).detach(), t.$slideTrack.detach(), t.$list.detach(), t.$slider.append(t.$slides)), t.cleanUpRows(), t.$slider.removeClass("slick-slider"), t.$slider.removeClass("slick-initialized"), t.$slider.removeClass("slick-dotted"), t.unslicked = !0, e || t.$slider.trigger("destroy", [t])
    }, e.prototype.disableTransition = function(i) {
        var e = this,
            t = {};
        t[e.transitionType] = "", !1 === e.options.fade ? e.$slideTrack.css(t) : e.$slides.eq(i).css(t)
    }, e.prototype.fadeSlide = function(i, e) {
        var t = this;
        !1 === t.cssTransitions ? (t.$slides.eq(i).css({
            zIndex: t.options.zIndex
        }), t.$slides.eq(i).animate({
            opacity: 1
        }, t.options.speed, t.options.easing, e)) : (t.applyTransition(i), t.$slides.eq(i).css({
            opacity: 1,
            zIndex: t.options.zIndex
        }), e && setTimeout(function() {
            t.disableTransition(i), e.call()
        }, t.options.speed))
    }, e.prototype.fadeSlideOut = function(i) {
        var e = this;
        !1 === e.cssTransitions ? e.$slides.eq(i).animate({
            opacity: 0,
            zIndex: e.options.zIndex - 2
        }, e.options.speed, e.options.easing) : (e.applyTransition(i), e.$slides.eq(i).css({
            opacity: 0,
            zIndex: e.options.zIndex - 2
        }))
    }, e.prototype.filterSlides = e.prototype.slickFilter = function(i) {
        var e = this;
        null !== i && (e.$slidesCache = e.$slides, e.unload(), e.$slideTrack.children(this.options.slide).detach(), e.$slidesCache.filter(i).appendTo(e.$slideTrack), e.reinit())
    }, e.prototype.focusHandler = function() {
        var e = this;
        e.$slider.off("focus.slick blur.slick").on("focus.slick blur.slick", "*", function(t) {
            t.stopImmediatePropagation();
            var o = i(this);
            setTimeout(function() {
                e.options.pauseOnFocus && (e.focussed = o.is(":focus"), e.autoPlay())
            }, 0)
        })
    }, e.prototype.getCurrent = e.prototype.slickCurrentSlide = function() {
        return this.currentSlide
    }, e.prototype.getDotCount = function() {
        var i = this,
            e = 0,
            t = 0,
            o = 0;
        if (!0 === i.options.infinite)
            if (i.slideCount <= i.options.slidesToShow) ++o;
            else
                for (; e < i.slideCount;) ++o, e = t + i.options.slidesToScroll, t += i.options.slidesToScroll <= i.options.slidesToShow ? i.options.slidesToScroll : i.options.slidesToShow;
        else if (!0 === i.options.centerMode) o = i.slideCount;
        else if (i.options.asNavFor)
            for (; e < i.slideCount;) ++o, e = t + i.options.slidesToScroll, t += i.options.slidesToScroll <= i.options.slidesToShow ? i.options.slidesToScroll : i.options.slidesToShow;
        else o = 1 + Math.ceil((i.slideCount - i.options.slidesToShow) / i.options.slidesToScroll);
        return o - 1
    }, e.prototype.getLeft = function(i) {
        var e, t, o, s, n = this,
            r = 0;
        return n.slideOffset = 0, t = n.$slides.first().outerHeight(!0), !0 === n.options.infinite ? (n.slideCount > n.options.slidesToShow && (n.slideOffset = n.slideWidth * n.options.slidesToShow * -1, s = -1, !0 === n.options.vertical && !0 === n.options.centerMode && (2 === n.options.slidesToShow ? s = -1.5 : 1 === n.options.slidesToShow && (s = -2)), r = t * n.options.slidesToShow * s), n.slideCount % n.options.slidesToScroll != 0 && i + n.options.slidesToScroll > n.slideCount && n.slideCount > n.options.slidesToShow && (i > n.slideCount ? (n.slideOffset = (n.options.slidesToShow - (i - n.slideCount)) * n.slideWidth * -1, r = (n.options.slidesToShow - (i - n.slideCount)) * t * -1) : (n.slideOffset = n.slideCount % n.options.slidesToScroll * n.slideWidth * -1, r = n.slideCount % n.options.slidesToScroll * t * -1))) : i + n.options.slidesToShow > n.slideCount && (n.slideOffset = (i + n.options.slidesToShow - n.slideCount) * n.slideWidth, r = (i + n.options.slidesToShow - n.slideCount) * t), n.slideCount <= n.options.slidesToShow && (n.slideOffset = 0, r = 0), !0 === n.options.centerMode && n.slideCount <= n.options.slidesToShow ? n.slideOffset = n.slideWidth * Math.floor(n.options.slidesToShow) / 2 - n.slideWidth * n.slideCount / 2 : !0 === n.options.centerMode && !0 === n.options.infinite ? n.slideOffset += n.slideWidth * Math.floor(n.options.slidesToShow / 2) - n.slideWidth : !0 === n.options.centerMode && (n.slideOffset = 0, n.slideOffset += n.slideWidth * Math.floor(n.options.slidesToShow / 2)), e = !1 === n.options.vertical ? i * n.slideWidth * -1 + n.slideOffset : i * t * -1 + r, !0 === n.options.variableWidth && (o = n.slideCount <= n.options.slidesToShow || !1 === n.options.infinite ? n.$slideTrack.children(".slick-slide").eq(i) : n.$slideTrack.children(".slick-slide").eq(i + n.options.slidesToShow), e = !0 === n.options.rtl ? o[0] ? -1 * (n.$slideTrack.width() - o[0].offsetLeft - o.width()) : 0 : o[0] ? -1 * o[0].offsetLeft : 0, !0 === n.options.centerMode && (o = n.slideCount <= n.options.slidesToShow || !1 === n.options.infinite ? n.$slideTrack.children(".slick-slide").eq(i) : n.$slideTrack.children(".slick-slide").eq(i + n.options.slidesToShow + 1), e = !0 === n.options.rtl ? o[0] ? -1 * (n.$slideTrack.width() - o[0].offsetLeft - o.width()) : 0 : o[0] ? -1 * o[0].offsetLeft : 0, e += (n.$list.width() - o.outerWidth()) / 2)), e
    }, e.prototype.getOption = e.prototype.slickGetOption = function(i) {
        return this.options[i]
    }, e.prototype.getNavigableIndexes = function() {
        var i, e = this,
            t = 0,
            o = 0,
            s = [];
        for (!1 === e.options.infinite ? i = e.slideCount : (t = -1 * e.options.slidesToScroll, o = -1 * e.options.slidesToScroll, i = 2 * e.slideCount); t < i;) s.push(t), t = o + e.options.slidesToScroll, o += e.options.slidesToScroll <= e.options.slidesToShow ? e.options.slidesToScroll : e.options.slidesToShow;
        return s
    }, e.prototype.getSlick = function() {
        return this
    }, e.prototype.getSlideCount = function() {
        var e, t, o = this;
        return t = !0 === o.options.centerMode ? o.slideWidth * Math.floor(o.options.slidesToShow / 2) : 0, !0 === o.options.swipeToSlide ? (o.$slideTrack.find(".slick-slide").each(function(s, n) {
            if (n.offsetLeft - t + i(n).outerWidth() / 2 > -1 * o.swipeLeft) return e = n, !1
        }), Math.abs(i(e).attr("data-slick-index") - o.currentSlide) || 1) : o.options.slidesToScroll
    }, e.prototype.goTo = e.prototype.slickGoTo = function(i, e) {
        this.changeSlide({
            data: {
                message: "index",
                index: parseInt(i)
            }
        }, e)
    }, e.prototype.init = function(e) {
        var t = this;
        i(t.$slider).hasClass("slick-initialized") || (i(t.$slider).addClass("slick-initialized"), t.buildRows(), t.buildOut(), t.setProps(), t.startLoad(), t.loadSlider(), t.initializeEvents(), t.updateArrows(), t.updateDots(), t.checkResponsive(!0), t.focusHandler()), e && t.$slider.trigger("init", [t]), !0 === t.options.accessibility && t.initADA(), t.options.autoplay && (t.paused = !1, t.autoPlay())
    }, e.prototype.initADA = function() {
        var e = this,
            t = Math.ceil(e.slideCount / e.options.slidesToShow),
            o = e.getNavigableIndexes().filter(function(i) {
                return i >= 0 && i < e.slideCount
            });
        e.$slides.add(e.$slideTrack.find(".slick-cloned")).attr({
            "aria-hidden": "true",
            tabindex: "-1"
        }).find("a, input, button, select").attr({
            tabindex: "-1"
        }), null !== e.$dots && (e.$slides.not(e.$slideTrack.find(".slick-cloned")).each(function(t) {
            var s = o.indexOf(t);
            i(this).attr({
                role: "tabpanel",
                id: "slick-slide" + e.instanceUid + t,
                tabindex: -1
            }), -1 !== s && i(this).attr({
                "aria-describedby": "slick-slide-control" + e.instanceUid + s
            })
        }), e.$dots.attr("role", "tablist").find("li").each(function(s) {
            var n = o[s];
            i(this).attr({
                role: "presentation"
            }), i(this).find("button").first().attr({
                role: "tab",
                id: "slick-slide-control" + e.instanceUid + s,
                "aria-controls": "slick-slide" + e.instanceUid + n,
                "aria-label": s + 1 + " of " + t,
                "aria-selected": null,
                tabindex: "-1"
            })
        }).eq(e.currentSlide).find("button").attr({
            "aria-selected": "true",
            tabindex: "0"
        }).end());
        for (var s = e.currentSlide, n = s + e.options.slidesToShow; s < n; s++) e.$slides.eq(s).attr("tabindex", 0);
        e.activateADA()
    }, e.prototype.initArrowEvents = function() {
        var i = this;
        !0 === i.options.arrows && i.slideCount > i.options.slidesToShow && (i.$prevArrow.off("click.slick").on("click.slick", {
            message: "previous"
        }, i.changeSlide), i.$nextArrow.off("click.slick").on("click.slick", {
            message: "next"
        }, i.changeSlide), !0 === i.options.accessibility && (i.$prevArrow.on("keydown.slick", i.keyHandler), i.$nextArrow.on("keydown.slick", i.keyHandler)))
    }, e.prototype.initDotEvents = function() {
        var e = this;
        !0 === e.options.dots && (i("li", e.$dots).on("click.slick", {
            message: "index"
        }, e.changeSlide), !0 === e.options.accessibility && e.$dots.on("keydown.slick", e.keyHandler)), !0 === e.options.dots && !0 === e.options.pauseOnDotsHover && i("li", e.$dots).on("mouseenter.slick", i.proxy(e.interrupt, e, !0)).on("mouseleave.slick", i.proxy(e.interrupt, e, !1))
    }, e.prototype.initSlideEvents = function() {
        var e = this;
        e.options.pauseOnHover && (e.$list.on("mouseenter.slick", i.proxy(e.interrupt, e, !0)), e.$list.on("mouseleave.slick", i.proxy(e.interrupt, e, !1)))
    }, e.prototype.initializeEvents = function() {
        var e = this;
        e.initArrowEvents(), e.initDotEvents(), e.initSlideEvents(), e.$list.on("touchstart.slick mousedown.slick", {
            action: "start"
        }, e.swipeHandler), e.$list.on("touchmove.slick mousemove.slick", {
            action: "move"
        }, e.swipeHandler), e.$list.on("touchend.slick mouseup.slick", {
            action: "end"
        }, e.swipeHandler), e.$list.on("touchcancel.slick mouseleave.slick", {
            action: "end"
        }, e.swipeHandler), e.$list.on("click.slick", e.clickHandler), i(document).on(e.visibilityChange, i.proxy(e.visibility, e)), !0 === e.options.accessibility && e.$list.on("keydown.slick", e.keyHandler), !0 === e.options.focusOnSelect && i(e.$slideTrack).children().on("click.slick", e.selectHandler), i(window).on("orientationchange.slick.slick-" + e.instanceUid, i.proxy(e.orientationChange, e)), i(window).on("resize.slick.slick-" + e.instanceUid, i.proxy(e.resize, e)), i("[draggable!=true]", e.$slideTrack).on("dragstart", e.preventDefault), i(window).on("load.slick.slick-" + e.instanceUid, e.setPosition), i(e.setPosition)
    }, e.prototype.initUI = function() {
        var i = this;
        !0 === i.options.arrows && i.slideCount > i.options.slidesToShow && (i.$prevArrow.show(), i.$nextArrow.show()), !0 === i.options.dots && i.slideCount > i.options.slidesToShow && i.$dots.show()
    }, e.prototype.keyHandler = function(i) {
        var e = this;
        i.target.tagName.match("TEXTAREA|INPUT|SELECT") || (37 === i.keyCode && !0 === e.options.accessibility ? e.changeSlide({
            data: {
                message: !0 === e.options.rtl ? "next" : "previous"
            }
        }) : 39 === i.keyCode && !0 === e.options.accessibility && e.changeSlide({
            data: {
                message: !0 === e.options.rtl ? "previous" : "next"
            }
        }))
    }, e.prototype.lazyLoad = function() {
        function e(e) {
            i("img[data-lazy]", e).each(function() {
                var e = i(this),
                    t = i(this).attr("data-lazy"),
                    o = i(this).attr("data-srcset"),
                    s = i(this).attr("data-sizes") || n.$slider.attr("data-sizes"),
                    r = document.createElement("img");
                r.onload = function() {
                    e.animate({
                        opacity: 0
                    }, 100, function() {
                        o && (e.attr("srcset", o), s && e.attr("sizes", s)), e.attr("src", t).animate({
                            opacity: 1
                        }, 200, function() {
                            e.removeAttr("data-lazy data-srcset data-sizes").removeClass("slick-loading")
                        }), n.$slider.trigger("lazyLoaded", [n, e, t])
                    })
                }, r.onerror = function() {
                    e.removeAttr("data-lazy").removeClass("slick-loading").addClass("slick-lazyload-error"), n.$slider.trigger("lazyLoadError", [n, e, t])
                }, r.src = t
            })
        }
        var t, o, s, n = this;
        if (!0 === n.options.centerMode ? !0 === n.options.infinite ? s = (o = n.currentSlide + (n.options.slidesToShow / 2 + 1)) + n.options.slidesToShow + 2 : (o = Math.max(0, n.currentSlide - (n.options.slidesToShow / 2 + 1)), s = n.options.slidesToShow / 2 + 1 + 2 + n.currentSlide) : (o = n.options.infinite ? n.options.slidesToShow + n.currentSlide : n.currentSlide, s = Math.ceil(o + n.options.slidesToShow), !0 === n.options.fade && (o > 0 && o--, s <= n.slideCount && s++)), t = n.$slider.find(".slick-slide").slice(o, s), "anticipated" === n.options.lazyLoad)
            for (var r = o - 1, l = s, d = n.$slider.find(".slick-slide"), a = 0; a < n.options.slidesToScroll; a++) r < 0 && (r = n.slideCount - 1), t = (t = t.add(d.eq(r))).add(d.eq(l)), r--, l++;
        e(t), n.slideCount <= n.options.slidesToShow ? e(n.$slider.find(".slick-slide")) : n.currentSlide >= n.slideCount - n.options.slidesToShow ? e(n.$slider.find(".slick-cloned").slice(0, n.options.slidesToShow)) : 0 === n.currentSlide && e(n.$slider.find(".slick-cloned").slice(-1 * n.options.slidesToShow))
    }, e.prototype.loadSlider = function() {
        var i = this;
        i.setPosition(), i.$slideTrack.css({
            opacity: 1
        }), i.$slider.removeClass("slick-loading"), i.initUI(), "progressive" === i.options.lazyLoad && i.progressiveLazyLoad()
    }, e.prototype.next = e.prototype.slickNext = function() {
        this.changeSlide({
            data: {
                message: "next"
            }
        })
    }, e.prototype.orientationChange = function() {
        var i = this;
        i.checkResponsive(), i.setPosition()
    }, e.prototype.pause = e.prototype.slickPause = function() {
        var i = this;
        i.autoPlayClear(), i.paused = !0
    }, e.prototype.play = e.prototype.slickPlay = function() {
        var i = this;
        i.autoPlay(), i.options.autoplay = !0, i.paused = !1, i.focussed = !1, i.interrupted = !1
    }, e.prototype.postSlide = function(e) {
        var t = this;
        t.unslicked || (t.$slider.trigger("afterChange", [t, e]), t.animating = !1, t.slideCount > t.options.slidesToShow && t.setPosition(), t.swipeLeft = null, t.options.autoplay && t.autoPlay(), !0 === t.options.accessibility && (t.initADA(), t.options.focusOnChange && i(t.$slides.get(t.currentSlide)).attr("tabindex", 0).focus()))
    }, e.prototype.prev = e.prototype.slickPrev = function() {
        this.changeSlide({
            data: {
                message: "previous"
            }
        })
    }, e.prototype.preventDefault = function(i) {
        i.preventDefault()
    }, e.prototype.progressiveLazyLoad = function(e) {
        e = e || 1;
        var t, o, s, n, r, l = this,
            d = i("img[data-lazy]", l.$slider);
        d.length ? (t = d.first(), o = t.attr("data-lazy"), s = t.attr("data-srcset"), n = t.attr("data-sizes") || l.$slider.attr("data-sizes"), (r = document.createElement("img")).onload = function() {
            s && (t.attr("srcset", s), n && t.attr("sizes", n)), t.attr("src", o).removeAttr("data-lazy data-srcset data-sizes").removeClass("slick-loading"), !0 === l.options.adaptiveHeight && l.setPosition(), l.$slider.trigger("lazyLoaded", [l, t, o]), l.progressiveLazyLoad()
        }, r.onerror = function() {
            e < 3 ? setTimeout(function() {
                l.progressiveLazyLoad(e + 1)
            }, 500) : (t.removeAttr("data-lazy").removeClass("slick-loading").addClass("slick-lazyload-error"), l.$slider.trigger("lazyLoadError", [l, t, o]), l.progressiveLazyLoad())
        }, r.src = o) : l.$slider.trigger("allImagesLoaded", [l])
    }, e.prototype.refresh = function(e) {
        var t, o, s = this;
        o = s.slideCount - s.options.slidesToShow, !s.options.infinite && s.currentSlide > o && (s.currentSlide = o), s.slideCount <= s.options.slidesToShow && (s.currentSlide = 0), t = s.currentSlide, s.destroy(!0), i.extend(s, s.initials, {
            currentSlide: t
        }), s.init(), e || s.changeSlide({
            data: {
                message: "index",
                index: t
            }
        }, !1)
    }, e.prototype.registerBreakpoints = function() {
        var e, t, o, s = this,
            n = s.options.responsive || null;
        if ("array" === i.type(n) && n.length) {
            s.respondTo = s.options.respondTo || "window";
            for (e in n)
                if (o = s.breakpoints.length - 1, n.hasOwnProperty(e)) {
                    for (t = n[e].breakpoint; o >= 0;) s.breakpoints[o] && s.breakpoints[o] === t && s.breakpoints.splice(o, 1), o--;
                    s.breakpoints.push(t), s.breakpointSettings[t] = n[e].settings
                }
            s.breakpoints.sort(function(i, e) {
                return s.options.mobileFirst ? i - e : e - i
            })
        }
    }, e.prototype.reinit = function() {
        var e = this;
        e.$slides = e.$slideTrack.children(e.options.slide).addClass("slick-slide"), e.slideCount = e.$slides.length, e.currentSlide >= e.slideCount && 0 !== e.currentSlide && (e.currentSlide = e.currentSlide - e.options.slidesToScroll), e.slideCount <= e.options.slidesToShow && (e.currentSlide = 0), e.registerBreakpoints(), e.setProps(), e.setupInfinite(), e.buildArrows(), e.updateArrows(), e.initArrowEvents(), e.buildDots(), e.updateDots(), e.initDotEvents(), e.cleanUpSlideEvents(), e.initSlideEvents(), e.checkResponsive(!1, !0), !0 === e.options.focusOnSelect && i(e.$slideTrack).children().on("click.slick", e.selectHandler), e.setSlideClasses("number" == typeof e.currentSlide ? e.currentSlide : 0), e.setPosition(), e.focusHandler(), e.paused = !e.options.autoplay, e.autoPlay(), e.$slider.trigger("reInit", [e])
    }, e.prototype.resize = function() {
        var e = this;
        i(window).width() !== e.windowWidth && (clearTimeout(e.windowDelay), e.windowDelay = window.setTimeout(function() {
            e.windowWidth = i(window).width(), e.checkResponsive(), e.unslicked || e.setPosition()
        }, 50))
    }, e.prototype.removeSlide = e.prototype.slickRemove = function(i, e, t) {
        var o = this;
        if (i = "boolean" == typeof i ? !0 === (e = i) ? 0 : o.slideCount - 1 : !0 === e ? --i : i, o.slideCount < 1 || i < 0 || i > o.slideCount - 1) return !1;
        o.unload(), !0 === t ? o.$slideTrack.children().remove() : o.$slideTrack.children(this.options.slide).eq(i).remove(), o.$slides = o.$slideTrack.children(this.options.slide), o.$slideTrack.children(this.options.slide).detach(), o.$slideTrack.append(o.$slides), o.$slidesCache = o.$slides, o.reinit()
    }, e.prototype.setCSS = function(i) {
        var e, t, o = this,
            s = {};
        !0 === o.options.rtl && (i = -i), e = "left" == o.positionProp ? Math.ceil(i) + "px" : "0px", t = "top" == o.positionProp ? Math.ceil(i) + "px" : "0px", s[o.positionProp] = i, !1 === o.transformsEnabled ? o.$slideTrack.css(s) : (s = {}, !1 === o.cssTransitions ? (s[o.animType] = "translate(" + e + ", " + t + ")", o.$slideTrack.css(s)) : (s[o.animType] = "translate3d(" + e + ", " + t + ", 0px)", o.$slideTrack.css(s)))
    }, e.prototype.setDimensions = function() {
        var i = this;
        !1 === i.options.vertical ? !0 === i.options.centerMode && i.$list.css({
            padding: "0px " + i.options.centerPadding
        }) : (i.$list.height(i.$slides.first().outerHeight(!0) * i.options.slidesToShow), !0 === i.options.centerMode && i.$list.css({
            padding: i.options.centerPadding + " 0px"
        })), i.listWidth = i.$list.width(), i.listHeight = i.$list.height(), !1 === i.options.vertical && !1 === i.options.variableWidth ? (i.slideWidth = Math.ceil(i.listWidth / i.options.slidesToShow), i.$slideTrack.width(Math.ceil(i.slideWidth * i.$slideTrack.children(".slick-slide").length))) : !0 === i.options.variableWidth ? i.$slideTrack.width(5e3 * i.slideCount) : (i.slideWidth = Math.ceil(i.listWidth), i.$slideTrack.height(Math.ceil(i.$slides.first().outerHeight(!0) * i.$slideTrack.children(".slick-slide").length)));
        var e = i.$slides.first().outerWidth(!0) - i.$slides.first().width();
        !1 === i.options.variableWidth && i.$slideTrack.children(".slick-slide").width(i.slideWidth - e)
    }, e.prototype.setFade = function() {
        var e, t = this;
        t.$slides.each(function(o, s) {
            e = t.slideWidth * o * -1, !0 === t.options.rtl ? i(s).css({
                position: "relative",
                right: e,
                top: 0,
                zIndex: t.options.zIndex - 2,
                opacity: 0
            }) : i(s).css({
                position: "relative",
                left: e,
                top: 0,
                zIndex: t.options.zIndex - 2,
                opacity: 0
            })
        }), t.$slides.eq(t.currentSlide).css({
            zIndex: t.options.zIndex - 1,
            opacity: 1
        })
    }, e.prototype.setHeight = function() {
        var i = this;
        if (1 === i.options.slidesToShow && !0 === i.options.adaptiveHeight && !1 === i.options.vertical) {
            var e = i.$slides.eq(i.currentSlide).outerHeight(!0);
            i.$list.css("height", e)
        }
    }, e.prototype.setOption = e.prototype.slickSetOption = function() {
        var e, t, o, s, n, r = this,
            l = !1;
        if ("object" === i.type(arguments[0]) ? (o = arguments[0], l = arguments[1], n = "multiple") : "string" === i.type(arguments[0]) && (o = arguments[0], s = arguments[1], l = arguments[2], "responsive" === arguments[0] && "array" === i.type(arguments[1]) ? n = "responsive" : void 0 !== arguments[1] && (n = "single")), "single" === n) r.options[o] = s;
        else if ("multiple" === n) i.each(o, function(i, e) {
            r.options[i] = e
        });
        else if ("responsive" === n)
            for (t in s)
                if ("array" !== i.type(r.options.responsive)) r.options.responsive = [s[t]];
                else {
                    for (e = r.options.responsive.length - 1; e >= 0;) r.options.responsive[e].breakpoint === s[t].breakpoint && r.options.responsive.splice(e, 1), e--;
                    r.options.responsive.push(s[t])
                }
        l && (r.unload(), r.reinit())
    }, e.prototype.setPosition = function() {
        var i = this;
        i.setDimensions(), i.setHeight(), !1 === i.options.fade ? i.setCSS(i.getLeft(i.currentSlide)) : i.setFade(), i.$slider.trigger("setPosition", [i])
    }, e.prototype.setProps = function() {
        var i = this,
            e = document.body.style;
        i.positionProp = !0 === i.options.vertical ? "top" : "left", "top" === i.positionProp ? i.$slider.addClass("slick-vertical") : i.$slider.removeClass("slick-vertical"), void 0 === e.WebkitTransition && void 0 === e.MozTransition && void 0 === e.msTransition || !0 === i.options.useCSS && (i.cssTransitions = !0), i.options.fade && ("number" == typeof i.options.zIndex ? i.options.zIndex < 3 && (i.options.zIndex = 3) : i.options.zIndex = i.defaults.zIndex), void 0 !== e.OTransform && (i.animType = "OTransform", i.transformType = "-o-transform", i.transitionType = "OTransition", void 0 === e.perspectiveProperty && void 0 === e.webkitPerspective && (i.animType = !1)), void 0 !== e.MozTransform && (i.animType = "MozTransform", i.transformType = "-moz-transform", i.transitionType = "MozTransition", void 0 === e.perspectiveProperty && void 0 === e.MozPerspective && (i.animType = !1)), void 0 !== e.webkitTransform && (i.animType = "webkitTransform", i.transformType = "-webkit-transform", i.transitionType = "webkitTransition", void 0 === e.perspectiveProperty && void 0 === e.webkitPerspective && (i.animType = !1)), void 0 !== e.msTransform && (i.animType = "msTransform", i.transformType = "-ms-transform", i.transitionType = "msTransition", void 0 === e.msTransform && (i.animType = !1)), void 0 !== e.transform && !1 !== i.animType && (i.animType = "transform", i.transformType = "transform", i.transitionType = "transition"), i.transformsEnabled = i.options.useTransform && null !== i.animType && !1 !== i.animType
    }, e.prototype.setSlideClasses = function(i) {
        var e, t, o, s, n = this;
        if (t = n.$slider.find(".slick-slide").removeClass("slick-active slick-center slick-current").attr("aria-hidden", "true"), n.$slides.eq(i).addClass("slick-current"), !0 === n.options.centerMode) {
            var r = n.options.slidesToShow % 2 == 0 ? 1 : 0;
            e = Math.floor(n.options.slidesToShow / 2), !0 === n.options.infinite && (i >= e && i <= n.slideCount - 1 - e ? n.$slides.slice(i - e + r, i + e + 1).addClass("slick-active").attr("aria-hidden", "false") : (o = n.options.slidesToShow + i, t.slice(o - e + 1 + r, o + e + 2).addClass("slick-active").attr("aria-hidden", "false")), 0 === i ? t.eq(t.length - 1 - n.options.slidesToShow).addClass("slick-center") : i === n.slideCount - 1 && t.eq(n.options.slidesToShow).addClass("slick-center")), n.$slides.eq(i).addClass("slick-center")
        } else i >= 0 && i <= n.slideCount - n.options.slidesToShow ? n.$slides.slice(i, i + n.options.slidesToShow).addClass("slick-active").attr("aria-hidden", "false") : t.length <= n.options.slidesToShow ? t.addClass("slick-active").attr("aria-hidden", "false") : (s = n.slideCount % n.options.slidesToShow, o = !0 === n.options.infinite ? n.options.slidesToShow + i : i, n.options.slidesToShow == n.options.slidesToScroll && n.slideCount - i < n.options.slidesToShow ? t.slice(o - (n.options.slidesToShow - s), o + s).addClass("slick-active").attr("aria-hidden", "false") : t.slice(o, o + n.options.slidesToShow).addClass("slick-active").attr("aria-hidden", "false"));
        "ondemand" !== n.options.lazyLoad && "anticipated" !== n.options.lazyLoad || n.lazyLoad()
    }, e.prototype.setupInfinite = function() {
        var e, t, o, s = this;
        if (!0 === s.options.fade && (s.options.centerMode = !1), !0 === s.options.infinite && !1 === s.options.fade && (t = null, s.slideCount > s.options.slidesToShow)) {
            for (o = !0 === s.options.centerMode ? s.options.slidesToShow + 1 : s.options.slidesToShow, e = s.slideCount; e > s.slideCount - o; e -= 1) t = e - 1, i(s.$slides[t]).clone(!0).attr("id", "").attr("data-slick-index", t - s.slideCount).prependTo(s.$slideTrack).addClass("slick-cloned");
            for (e = 0; e < o + s.slideCount; e += 1) t = e, i(s.$slides[t]).clone(!0).attr("id", "").attr("data-slick-index", t + s.slideCount).appendTo(s.$slideTrack).addClass("slick-cloned");
            s.$slideTrack.find(".slick-cloned").find("[id]").each(function() {
                i(this).attr("id", "")
            })
        }
    }, e.prototype.interrupt = function(i) {
        var e = this;
        i || e.autoPlay(), e.interrupted = i
    }, e.prototype.selectHandler = function(e) {
        var t = this,
            o = i(e.target).is(".slick-slide") ? i(e.target) : i(e.target).parents(".slick-slide"),
            s = parseInt(o.attr("data-slick-index"));
        s || (s = 0), t.slideCount <= t.options.slidesToShow ? t.slideHandler(s, !1, !0) : t.slideHandler(s)
    }, e.prototype.slideHandler = function(i, e, t) {
        var o, s, n, r, l, d = null,
            a = this;
        if (e = e || !1, !(!0 === a.animating && !0 === a.options.waitForAnimate || !0 === a.options.fade && a.currentSlide === i))
            if (!1 === e && a.asNavFor(i), o = i, d = a.getLeft(o), r = a.getLeft(a.currentSlide), a.currentLeft = null === a.swipeLeft ? r : a.swipeLeft, !1 === a.options.infinite && !1 === a.options.centerMode && (i < 0 || i > a.getDotCount() * a.options.slidesToScroll)) !1 === a.options.fade && (o = a.currentSlide, !0 !== t ? a.animateSlide(r, function() {
                a.postSlide(o)
            }) : a.postSlide(o));
            else if (!1 === a.options.infinite && !0 === a.options.centerMode && (i < 0 || i > a.slideCount - a.options.slidesToScroll)) !1 === a.options.fade && (o = a.currentSlide, !0 !== t ? a.animateSlide(r, function() {
            a.postSlide(o)
        }) : a.postSlide(o));
        else {
            if (a.options.autoplay && clearInterval(a.autoPlayTimer), s = o < 0 ? a.slideCount % a.options.slidesToScroll != 0 ? a.slideCount - a.slideCount % a.options.slidesToScroll : a.slideCount + o : o >= a.slideCount ? a.slideCount % a.options.slidesToScroll != 0 ? 0 : o - a.slideCount : o, a.animating = !0, a.$slider.trigger("beforeChange", [a, a.currentSlide, s]), n = a.currentSlide, a.currentSlide = s, a.setSlideClasses(a.currentSlide), a.options.asNavFor && (l = (l = a.getNavTarget()).slick("getSlick")).slideCount <= l.options.slidesToShow && l.setSlideClasses(a.currentSlide), a.updateDots(), a.updateArrows(), !0 === a.options.fade) return !0 !== t ? (a.fadeSlideOut(n), a.fadeSlide(s, function() {
                a.postSlide(s)
            })) : a.postSlide(s), void a.animateHeight();
            !0 !== t ? a.animateSlide(d, function() {
                a.postSlide(s)
            }) : a.postSlide(s)
        }
    }, e.prototype.startLoad = function() {
        var i = this;
        !0 === i.options.arrows && i.slideCount > i.options.slidesToShow && (i.$prevArrow.hide(), i.$nextArrow.hide()), !0 === i.options.dots && i.slideCount > i.options.slidesToShow && i.$dots.hide(), i.$slider.addClass("slick-loading")
    }, e.prototype.swipeDirection = function() {
        var i, e, t, o, s = this;
        return i = s.touchObject.startX - s.touchObject.curX, e = s.touchObject.startY - s.touchObject.curY, t = Math.atan2(e, i), (o = Math.round(180 * t / Math.PI)) < 0 && (o = 360 - Math.abs(o)), o <= 45 && o >= 0 ? !1 === s.options.rtl ? "left" : "right" : o <= 360 && o >= 315 ? !1 === s.options.rtl ? "left" : "right" : o >= 135 && o <= 225 ? !1 === s.options.rtl ? "right" : "left" : !0 === s.options.verticalSwiping ? o >= 35 && o <= 135 ? "down" : "up" : "vertical"
    }, e.prototype.swipeEnd = function(i) {
        var e, t, o = this;
        if (o.dragging = !1, o.swiping = !1, o.scrolling) return o.scrolling = !1, !1;
        if (o.interrupted = !1, o.shouldClick = !(o.touchObject.swipeLength > 10), void 0 === o.touchObject.curX) return !1;
        if (!0 === o.touchObject.edgeHit && o.$slider.trigger("edge", [o, o.swipeDirection()]), o.touchObject.swipeLength >= o.touchObject.minSwipe) {
            switch (t = o.swipeDirection()) {
                case "left":
                case "down":
                    e = o.options.swipeToSlide ? o.checkNavigable(o.currentSlide + o.getSlideCount()) : o.currentSlide + o.getSlideCount(), o.currentDirection = 0;
                    break;
                case "right":
                case "up":
                    e = o.options.swipeToSlide ? o.checkNavigable(o.currentSlide - o.getSlideCount()) : o.currentSlide - o.getSlideCount(), o.currentDirection = 1
            }
            "vertical" != t && (o.slideHandler(e), o.touchObject = {}, o.$slider.trigger("swipe", [o, t]))
        } else o.touchObject.startX !== o.touchObject.curX && (o.slideHandler(o.currentSlide), o.touchObject = {})
    }, e.prototype.swipeHandler = function(i) {
        var e = this;
        if (!(!1 === e.options.swipe || "ontouchend" in document && !1 === e.options.swipe || !1 === e.options.draggable && -1 !== i.type.indexOf("mouse"))) switch (e.touchObject.fingerCount = i.originalEvent && void 0 !== i.originalEvent.touches ? i.originalEvent.touches.length : 1, e.touchObject.minSwipe = e.listWidth / e.options.touchThreshold, !0 === e.options.verticalSwiping && (e.touchObject.minSwipe = e.listHeight / e.options.touchThreshold), i.data.action) {
            case "start":
                e.swipeStart(i);
                break;
            case "move":
                e.swipeMove(i);
                break;
            case "end":
                e.swipeEnd(i)
        }
    }, e.prototype.swipeMove = function(i) {
        var e, t, o, s, n, r, l = this;
        return n = void 0 !== i.originalEvent ? i.originalEvent.touches : null, !(!l.dragging || l.scrolling || n && 1 !== n.length) && (e = l.getLeft(l.currentSlide), l.touchObject.curX = void 0 !== n ? n[0].pageX : i.clientX, l.touchObject.curY = void 0 !== n ? n[0].pageY : i.clientY, l.touchObject.swipeLength = Math.round(Math.sqrt(Math.pow(l.touchObject.curX - l.touchObject.startX, 2))), r = Math.round(Math.sqrt(Math.pow(l.touchObject.curY - l.touchObject.startY, 2))), !l.options.verticalSwiping && !l.swiping && r > 4 ? (l.scrolling = !0, !1) : (!0 === l.options.verticalSwiping && (l.touchObject.swipeLength = r), t = l.swipeDirection(), void 0 !== i.originalEvent && l.touchObject.swipeLength > 4 && (l.swiping = !0, i.preventDefault()), s = (!1 === l.options.rtl ? 1 : -1) * (l.touchObject.curX > l.touchObject.startX ? 1 : -1), !0 === l.options.verticalSwiping && (s = l.touchObject.curY > l.touchObject.startY ? 1 : -1), o = l.touchObject.swipeLength, l.touchObject.edgeHit = !1, !1 === l.options.infinite && (0 === l.currentSlide && "right" === t || l.currentSlide >= l.getDotCount() && "left" === t) && (o = l.touchObject.swipeLength * l.options.edgeFriction, l.touchObject.edgeHit = !0), !1 === l.options.vertical ? l.swipeLeft = e + o * s : l.swipeLeft = e + o * (l.$list.height() / l.listWidth) * s, !0 === l.options.verticalSwiping && (l.swipeLeft = e + o * s), !0 !== l.options.fade && !1 !== l.options.touchMove && (!0 === l.animating ? (l.swipeLeft = null, !1) : void l.setCSS(l.swipeLeft))))
    }, e.prototype.swipeStart = function(i) {
        var e, t = this;
        if (t.interrupted = !0, 1 !== t.touchObject.fingerCount || t.slideCount <= t.options.slidesToShow) return t.touchObject = {}, !1;
        void 0 !== i.originalEvent && void 0 !== i.originalEvent.touches && (e = i.originalEvent.touches[0]), t.touchObject.startX = t.touchObject.curX = void 0 !== e ? e.pageX : i.clientX, t.touchObject.startY = t.touchObject.curY = void 0 !== e ? e.pageY : i.clientY, t.dragging = !0
    }, e.prototype.unfilterSlides = e.prototype.slickUnfilter = function() {
        var i = this;
        null !== i.$slidesCache && (i.unload(), i.$slideTrack.children(this.options.slide).detach(), i.$slidesCache.appendTo(i.$slideTrack), i.reinit())
    }, e.prototype.unload = function() {
        var e = this;
        i(".slick-cloned", e.$slider).remove(), e.$dots && e.$dots.remove(), e.$prevArrow && e.htmlExpr.test(e.options.prevArrow) && e.$prevArrow.remove(), e.$nextArrow && e.htmlExpr.test(e.options.nextArrow) && e.$nextArrow.remove(), e.$slides.removeClass("slick-slide slick-active slick-visible slick-current").attr("aria-hidden", "true").css("width", "")
    }, e.prototype.unslick = function(i) {
        var e = this;
        e.$slider.trigger("unslick", [e, i]), e.destroy()
    }, e.prototype.updateArrows = function() {
        var i = this;
        Math.floor(i.options.slidesToShow / 2), !0 === i.options.arrows && i.slideCount > i.options.slidesToShow && !i.options.infinite && (i.$prevArrow.removeClass("slick-disabled").attr("aria-disabled", "false"), i.$nextArrow.removeClass("slick-disabled").attr("aria-disabled", "false"), 0 === i.currentSlide ? (i.$prevArrow.addClass("slick-disabled").attr("aria-disabled", "true"), i.$nextArrow.removeClass("slick-disabled").attr("aria-disabled", "false")) : i.currentSlide >= i.slideCount - i.options.slidesToShow && !1 === i.options.centerMode ? (i.$nextArrow.addClass("slick-disabled").attr("aria-disabled", "true"), i.$prevArrow.removeClass("slick-disabled").attr("aria-disabled", "false")) : i.currentSlide >= i.slideCount - 1 && !0 === i.options.centerMode && (i.$nextArrow.addClass("slick-disabled").attr("aria-disabled", "true"), i.$prevArrow.removeClass("slick-disabled").attr("aria-disabled", "false")))
    }, e.prototype.updateDots = function() {
        var i = this;
        null !== i.$dots && (i.$dots.find("li").removeClass("slick-active").end(), i.$dots.find("li").eq(Math.floor(i.currentSlide / i.options.slidesToScroll)).addClass("slick-active"))
    }, e.prototype.visibility = function() {
        var i = this;
        i.options.autoplay && (document[i.hidden] ? i.interrupted = !0 : i.interrupted = !1)
    }, i.fn.slick = function() {
        var i, t, o = this,
            s = arguments[0],
            n = Array.prototype.slice.call(arguments, 1),
            r = o.length;
        for (i = 0; i < r; i++)
            if ("object" == typeof s || void 0 === s ? o[i].slick = new e(o[i], s) : t = o[i].slick[s].apply(o[i].slick, n), void 0 !== t) return t;
        return o
    }
});