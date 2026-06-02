(function () {
    function installModalPlugin($) {
        if ($.fn && $.fn.modal) return;
        $.fn.modal = function (action) {
            return this.each(function () {
                if (action === 'show') {
                    this.style.display = 'block';
                    this.classList.add('show');
                    this.setAttribute('aria-modal', 'true');
                    document.body.classList.add('modal-open');
                } else if (action === 'hide') {
                    this.style.display = 'none';
                    this.classList.remove('show');
                    this.removeAttribute('aria-modal');
                    document.body.classList.remove('modal-open');
                    this.dispatchEvent(new Event('hidden.bs.modal', { bubbles: true }));
                }
            });
        };
    }

    if (window.jQuery && window.$) {
        installModalPlugin(window.$);
        return;
    }

    var dataStore = new WeakMap();

    function MiniQuery(items) {
        this.items = items || [];
        this.length = this.items.length;
        for (var i = 0; i < this.items.length; i++) this[i] = this.items[i];
    }

    function toArray(input) {
        if (!input) return [];
        if (input instanceof MiniQuery) return input.items;
        if (input instanceof Node || input === window || input === document) return [input];
        if (input instanceof NodeList || Array.isArray(input)) return Array.prototype.slice.call(input);
        return [];
    }

    function makeNodes(html) {
        var tpl = document.createElement('template');
        tpl.innerHTML = html.trim();
        return Array.prototype.slice.call(tpl.content.childNodes);
    }

    function $(selector) {
        if (typeof selector === 'function') {
            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', selector);
            } else {
                selector();
            }
            return new MiniQuery([]);
        }
        if (typeof selector === 'string') {
            if (selector.trim().charAt(0) === '<') return new MiniQuery(makeNodes(selector));
            return new MiniQuery(Array.prototype.slice.call(document.querySelectorAll(selector)));
        }
        return new MiniQuery(toArray(selector));
    }

    MiniQuery.prototype.each = function (callback) {
        this.items.forEach(function (el, idx) {
            callback.call(el, idx, el);
        });
        return this;
    };

    MiniQuery.prototype.on = function (eventName, selector, handler) {
        if (typeof selector === 'function') {
            handler = selector;
            selector = null;
        }
        return this.each(function () {
            this.addEventListener(eventName, function (event) {
                if (!selector) {
                    handler.call(event.currentTarget, event);
                    return;
                }
                var target = event.target.closest(selector);
                if (target && event.currentTarget.contains(target)) {
                    handler.call(target, event);
                }
            });
        });
    };

    MiniQuery.prototype.val = function (value) {
        if (value === undefined) return this.items[0] ? this.items[0].value : undefined;
        return this.each(function () { this.value = value; });
    };

    MiniQuery.prototype.text = function (value) {
        if (value === undefined) return this.items[0] ? this.items[0].textContent : undefined;
        return this.each(function () { this.textContent = value; });
    };

    MiniQuery.prototype.html = function (value) {
        if (value === undefined) return this.items[0] ? this.items[0].innerHTML : undefined;
        return this.each(function () { this.innerHTML = value; });
    };

    MiniQuery.prototype.empty = function () {
        return this.html('');
    };

    MiniQuery.prototype.append = function () {
        var args = Array.prototype.slice.call(arguments);
        return this.each(function () {
            var parent = this;
            args.forEach(function (arg) {
                if (typeof arg === 'string') {
                    makeNodes(arg).forEach(function (node) { parent.appendChild(node); });
                } else {
                    toArray(arg).forEach(function (node) { parent.appendChild(node); });
                }
            });
        });
    };

    MiniQuery.prototype.next = function (selector) {
        var el = this.items[0] ? this.items[0].nextElementSibling : null;
        if (el && selector && !el.matches(selector)) return new MiniQuery([]);
        return new MiniQuery(el ? [el] : []);
    };

    MiniQuery.prototype.find = function (selector) {
        var found = [];
        this.each(function () {
            found = found.concat(Array.prototype.slice.call(this.querySelectorAll(selector)));
        });
        return new MiniQuery(found);
    };

    MiniQuery.prototype.children = function (selector) {
        var found = [];
        this.each(function () {
            var kids = Array.prototype.slice.call(this.children);
            if (selector) kids = kids.filter(function (el) { return el.matches(selector); });
            found = found.concat(kids);
        });
        return new MiniQuery(found);
    };

    MiniQuery.prototype.eq = function (idx) {
        return new MiniQuery(this.items[idx] ? [this.items[idx]] : []);
    };

    MiniQuery.prototype.addClass = function (classes) {
        if (!classes) return this;
        var names = classes.split(/\s+/).filter(Boolean);
        return this.each(function () { this.classList.add.apply(this.classList, names); });
    };

    MiniQuery.prototype.removeClass = function (classes) {
        if (!classes) return this.each(function () { this.removeAttribute('class'); });
        var names = classes.split(/\s+/).filter(Boolean);
        return this.each(function () { this.classList.remove.apply(this.classList, names); });
    };

    MiniQuery.prototype.prop = function (name, value) {
        if (value === undefined) return this.items[0] ? this.items[0][name] : undefined;
        return this.each(function () { this[name] = value; });
    };

    MiniQuery.prototype.attr = function (name, value) {
        if (value === undefined) return this.items[0] ? this.items[0].getAttribute(name) : undefined;
        return this.each(function () { this.setAttribute(name, value); });
    };

    MiniQuery.prototype.data = function (name, value) {
        var attrName = name.replace(/[A-Z]/g, function (m) { return '-' + m.toLowerCase(); });
        if (value === undefined) {
            var el = this.items[0];
            if (!el) return undefined;
            var map = dataStore.get(el) || {};
            if (Object.prototype.hasOwnProperty.call(map, name)) return map[name];
            return el.dataset ? el.dataset[name] : el.getAttribute('data-' + attrName);
        }
        return this.each(function () {
            var map = dataStore.get(this) || {};
            map[name] = value;
            dataStore.set(this, map);
            if (this.dataset) this.dataset[name] = value;
        });
    };

    MiniQuery.prototype.show = function () {
        return this.each(function () { this.style.display = ''; });
    };

    MiniQuery.prototype.hide = function () {
        return this.each(function () { this.style.display = 'none'; });
    };

    MiniQuery.prototype.slideDown = MiniQuery.prototype.show;
    MiniQuery.prototype.slideUp = MiniQuery.prototype.hide;

    MiniQuery.prototype.is = function (selector) {
        var el = this.items[0];
        if (!el) return false;
        if (selector === ':visible') return !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
        return el.matches(selector);
    };

    MiniQuery.prototype.trigger = function (eventName) {
        return this.each(function () {
            if (eventName === 'focus' && this.focus) this.focus();
            if (eventName === 'select' && this.select) this.select();
            this.dispatchEvent(new Event(eventName, { bubbles: true }));
        });
    };

    MiniQuery.prototype.modal = function (action) {
        return this.each(function () {
            if (action === 'show') {
                this.style.display = 'block';
                this.classList.add('show');
                this.setAttribute('aria-modal', 'true');
                document.body.classList.add('modal-open');
            } else if (action === 'hide') {
                this.style.display = 'none';
                this.classList.remove('show');
                this.removeAttribute('aria-modal');
                document.body.classList.remove('modal-open');
                this.dispatchEvent(new Event('hidden.bs.modal', { bubbles: true }));
            }
        });
    };

    function ajax(options) {
        var method = options.type || options.method || 'GET';
        var headers = {};
        var body = options.data;
        if (options.contentType && options.contentType !== false) headers['Content-Type'] = options.contentType;

        var controller = null;
        var timer = null;
        if (options.timeout) {
            controller = new AbortController();
            timer = setTimeout(function () { controller.abort(); }, options.timeout);
        }

        fetch(options.url, {
            method: method,
            headers: headers,
            body: method.toUpperCase() === 'GET' ? undefined : body,
            signal: controller ? controller.signal : undefined
        }).then(function (res) {
            return res.text().then(function (text) {
                var json = null;
                try { json = text ? JSON.parse(text) : null; } catch (e) {}
                if (!res.ok) {
                    var xhr = { status: res.status, statusText: res.statusText, responseText: text };
                    if (options.error) options.error(xhr);
                    return;
                }
                if (options.success) options.success(json);
            });
        }).catch(function (err) {
            var xhr = { status: 0, statusText: err.name === 'AbortError' ? 'timeout' : 'error', responseText: '' };
            if (options.error) options.error(xhr);
        }).finally(function () {
            if (timer) clearTimeout(timer);
            if (options.complete) options.complete();
        });
    }

    $.ajax = ajax;
    $.getJSON = function (url, params, success) {
        if (typeof params === 'function') {
            success = params;
            params = null;
        }
        if (params) {
            var qs = new URLSearchParams(params).toString();
            url += (url.indexOf('?') === -1 ? '?' : '&') + qs;
        }
        var failHandler = null;
        ajax({
            url: url,
            type: 'GET',
            success: success,
            error: function (xhr) { if (failHandler) failHandler(xhr); }
        });
        return { fail: function (fn) { failHandler = fn; return this; } };
    };

    $.fn = MiniQuery.prototype;
    installModalPlugin($);
    window.$ = window.jQuery = $;

    document.addEventListener('click', function (event) {
        var closer = event.target.closest('[data-dismiss="modal"]');
        if (!closer) return;
        var modal = closer.closest('.modal');
        if (modal) $(modal).modal('hide');
    });
})();
