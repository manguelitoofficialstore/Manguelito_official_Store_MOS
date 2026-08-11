(() => {
  'use strict';

  const EXCEL_FILE = 'Productos.xlsx';
  const IMAGE_FOLDER = 'Imagenes_Finales';
  const POLL_MS = 60_000;
  const COLLECTION_ORDER = ['Oficial', 'Basic'];
  const CATEGORY_ORDER = ['Ropa', 'Accesorios', 'Items'];

  const state = {
    products: [],
    salesNumbers: [],
    cart: loadCart(),
    selectedProduct: null,
    selectedVariantId: '',
    detailQuantity: 1,
    lastFingerprint: ''
  };

  const $ = id => document.getElementById(id);
  const elements = {
    catalogRoot: $('catalogRoot'),
    emptyState: $('emptyState'),
    syncStatus: $('syncStatus'),
    cartButton: $('cartButton'),
    cartCount: $('cartCount'),
    cartDrawer: $('cartDrawer'),
    cartItems: $('cartItems'),
    cartTotal: $('cartTotal'),
    checkoutButton: $('checkoutButton'),
    productDialog: $('productDialog'),
    detailImage: $('detailImage'),
    detailImageFallback: $('detailImageFallback'),
    detailTitle: $('detailTitle'),
    detailReference: $('detailReference'),
    detailPrice: $('detailPrice'),
    detailDescription: $('detailDescription'),
    detailColor: $('detailColor'),
    detailSizeText: $('detailSizeText'),
    detailGender: $('detailGender'),
    detailStockNotice: $('detailStockNotice'),
    detailSizeSelect: $('detailSizeSelect'),
    detailQtyDec: $('detailQtyDec'),
    detailQtyInc: $('detailQtyInc'),
    detailQtyValue: $('detailQtyValue'),
    detailSubtotal: $('detailSubtotal'),
    detailError: $('detailError'),
    detailAddButton: $('detailAddButton'),
    checkoutDialog: $('checkoutDialog'),
    checkoutForm: $('checkoutForm'),
    zoneType: $('zoneType'),
    urbanAddressField: $('urbanAddressField'),
    urbanAddress: $('urbanAddress'),
    ruralPlaceField: $('ruralPlaceField'),
    ruralPlace: $('ruralPlace'),
    ruralReferenceField: $('ruralReferenceField'),
    ruralReference: $('ruralReference'),
    salesWhatsapp: $('salesWhatsapp'),
    checkoutError: $('checkoutError'),
    reportDialog: $('reportDialog'),
    reportTitle: $('reportTitle'),
    reportText: $('reportText'),
    copyReportButton: $('copyReportButton'),
    openWhatsappButton: $('openWhatsappButton')
  };

  function normalizeKey(value) {
    return String(value ?? '')
      .trim()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');
  }

  function rowToObject(row) {
    const out = {};
    Object.entries(row).forEach(([key, value]) => { out[normalizeKey(key)] = value; });
    return out;
  }

  function cleanText(value) { return String(value ?? '').trim(); }

  function parseNumber(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const text = cleanText(value).replace(/\s/g, '').replace(/\$/g, '');
    if (!text) return NaN;
    let normalized = text;
    if (/^-?\d{1,3}(\.\d{3})+$/.test(normalized)) normalized = normalized.replace(/\./g, '');
    else if (/^-?\d{1,3}(,\d{3})+$/.test(normalized)) normalized = normalized.replace(/,/g, '');
    else if (normalized.includes(',') && normalized.includes('.')) normalized = normalized.replace(/\./g, '').replace(',', '.');
    else normalized = normalized.replace(',', '.');
    return Number(normalized.replace(/[^0-9.-]/g, ''));
  }

  function normalizeCollection(value) {
    const v = normalizeKey(value);
    if (v.includes('basic')) return 'Basic';
    if (v.includes('oficial')) return 'Oficial';
    return cleanText(value);
  }

  function normalizeCategory(value) {
    const v = normalizeKey(value);
    if (v.startsWith('ropa')) return 'Ropa';
    if (v.startsWith('acces')) return 'Accesorios';
    if (v.startsWith('item')) return 'Items';
    return cleanText(value);
  }

  function makeVariantId(product) {
    return [product.referencia, product.size, product.color, product.gender, product.photo]
      .map(normalizeKey)
      .join('__');
  }

  function parseProduct(row) {
    const r = rowToObject(row);
    const referencia = cleanText(r.referencia ?? r.ref);
    const collectionRaw = cleanText(r.coleccion);
    const categoryRaw = cleanText(r.categoria);
    const priceRaw = r.preciodeventa ?? r.precioventa ?? r.precio;

    if (!referencia || !collectionRaw || !categoryRaw || priceRaw === '' || priceRaw == null) return null;

    const price = parseNumber(priceRaw);
    if (!Number.isFinite(price)) return null;

    const stockRaw = parseNumber(r.inventario ?? r.stock ?? 0);
    const stock = Math.max(0, Math.floor(Number.isFinite(stockRaw) ? stockRaw : 0));
    const descriptionBase = cleanText(r.descripcion);
    const sizeRaw = cleanText(r.talla);
    const size = sizeRaw || 'Única';
    const color = cleanText(r.color);
    const gender = cleanText(r.genero);
    const description = [descriptionBase, sizeRaw, color, gender].filter(Boolean).join(' · ');
    const subcategory = cleanText(r.subcategoria);
    const photo = cleanText(r.foto);

    const product = {
      referencia,
      collection: normalizeCollection(collectionRaw),
      category: normalizeCategory(categoryRaw),
      subcategory,
      name: descriptionBase || subcategory || referencia,
      descriptionBase: descriptionBase || referencia,
      description: description || descriptionBase || referencia,
      price,
      stock,
      size,
      color,
      gender,
      photo
    };
    product.variantId = makeVariantId(product);
    return product;
  }

  function normalizeSalesPhone(value) {
    let digits = cleanText(value).replace(/\D/g, '');
    if (!digits) return '';
    if (digits.length === 10 && digits.startsWith('3')) digits = `57${digits}`;
    return digits;
  }

  async function loadWorkbook({ silent = false } = {}) {
    if (!window.XLSX) {
      if (!silent) elements.syncStatus.textContent = 'No se pudo cargar el lector de Excel.';
      return;
    }

    try {
      const response = await fetch(`${EXCEL_FILE}?v=${Date.now()}`, { cache: 'no-store' });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const bytes = await response.arrayBuffer();
      const workbook = XLSX.read(bytes, { type: 'array' });

      const productSheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!productSheet) throw new Error('No existe la hoja de productos.');
      const productRows = XLSX.utils.sheet_to_json(productSheet, { defval: '' });
      const products = productRows.map(parseProduct).filter(Boolean);

      const salesNumbers = [];
      const configSheet = workbook.Sheets[workbook.SheetNames[1]];
      if (configSheet) {
        const rows = XLSX.utils.sheet_to_json(configSheet, { header: 1, defval: '' });
        rows.forEach(row => {
          const phone = normalizeSalesPhone(row[0]);
          if (phone) salesNumbers.push(phone);
        });
      }

      const fingerprint = JSON.stringify(products.map(p => [
        p.variantId, p.referencia, p.description, p.price, p.stock, p.photo, p.collection, p.category
      ])) + salesNumbers.join('|');
      const changed = fingerprint !== state.lastFingerprint;

      state.lastFingerprint = fingerprint;
      state.products = products;
      state.salesNumbers = [...new Set(salesNumbers)];
      pruneInvalidCart();

      if (changed || !silent) {
        renderCatalog();
        renderCart();
        renderSalesNumbers();
      }

      elements.syncStatus.textContent = `Excel sincronizado · ${products.length} producto${products.length === 1 ? '' : 's'} · ${new Date().toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' })}`;
    } catch (error) {
      console.error(error);
      elements.syncStatus.textContent = 'No se pudo leer Productos.xlsx.';
      if (!silent) {
        elements.catalogRoot.innerHTML = '';
        elements.emptyState.hidden = false;
        elements.emptyState.textContent = 'No fue posible cargar el catálogo. Abre la tienda desde GitHub Pages o un servidor web.';
      }
    }
  }

  function renderCatalog() {
    elements.catalogRoot.innerHTML = '';
    let rendered = 0;

    COLLECTION_ORDER.forEach(collection => {
      const collectionProducts = state.products.filter(p => p.collection === collection);
      if (!collectionProducts.length) return;

      const section = document.createElement('section');
      section.className = `collection-block collection-${collection.toLowerCase()}`;
      const logo = collection === 'Oficial' ? 'assets/logos/mos-oficial.png' : 'assets/logos/mb.png';
      const logoAlt = collection === 'Oficial' ? 'Logo oficial MOS' : 'Logo MosBasic';
      section.innerHTML = `
        <header class="collection-heading">
          <div class="collection-heading-copy">
            <p class="eyebrow">${collection === 'Oficial' ? 'COLECCIÓN' : 'MOS BASIC'}</p>
            <h2>${collection === 'Oficial' ? 'Productos oficiales' : 'Colección Basic'}</h2>
          </div>
          <div class="collection-heading-logo-wrap">
            <img class="collection-heading-logo" src="${logo}" alt="${logoAlt}" />
          </div>
        </header>`;

      CATEGORY_ORDER.forEach(category => {
        const categoryProducts = collectionProducts.filter(p => p.category === category);
        if (!categoryProducts.length) return;

        const block = document.createElement('section');
        block.className = 'category-block';
        block.innerHTML = `
          <div class="category-heading">
            <h3>${category}</h3>
            <span>${categoryProducts.length} producto${categoryProducts.length === 1 ? '' : 's'}</span>
          </div>
          <div class="product-grid"></div>`;

        const grid = block.querySelector('.product-grid');
        categoryProducts.forEach(product => {
          grid.appendChild(productCard(product));
          rendered += 1;
        });
        section.appendChild(block);
      });

      elements.catalogRoot.appendChild(section);
    });

    elements.emptyState.hidden = rendered > 0;
  }

  function photoCandidates(photo) {
    const name = cleanText(photo);
    if (!name) return [];

    if (/^(https?:)?\/\//i.test(name) || name.startsWith('./') || name.startsWith('../')) return [name];

    if (/\.[a-z0-9]{2,5}$/i.test(name)) return [`${IMAGE_FOLDER}/${name}`];

    return [
      `${IMAGE_FOLDER}/${name}`,
      `${IMAGE_FOLDER}/${name}.jpg`,
      `${IMAGE_FOLDER}/${name}.jpeg`,
      `${IMAGE_FOLDER}/${name}.png`,
      `${IMAGE_FOLDER}/${name}.webp`,
      `${IMAGE_FOLDER}/${name}.JPG`,
      `${IMAGE_FOLDER}/${name}.JPEG`,
      `${IMAGE_FOLDER}/${name}.PNG`,
      `${IMAGE_FOLDER}/${name}.WEBP`
    ];
  }

  function attachPhotoFallback(img, fallback, product) {
    const candidates = photoCandidates(product.photo);
    let index = 0;

    img.onload = () => {
      img.hidden = false;
      fallback.hidden = true;
    };

    img.onerror = () => {
      index += 1;
      if (index < candidates.length) {
        img.src = candidates[index];
      } else {
        img.hidden = true;
        fallback.hidden = false;
      }
    };

    if (!candidates.length) {
      img.hidden = true;
      fallback.hidden = false;
      return;
    }

    img.hidden = false;
    fallback.hidden = true;
    img.src = candidates[index];
  }

  function productCard(product) {
    const article = document.createElement('article');
    article.className = 'product-card';
    article.innerHTML = `
      <div class="product-image-wrap">
        <img class="product-image" alt="${escapeHtml(product.description)}" loading="lazy" />
        <div class="product-image-fallback" hidden></div>
      </div>
      <div class="product-body">
        <div class="product-meta">
          <span>${escapeHtml(product.subcategory || product.category)}</span>
          <span>${escapeHtml(product.referencia)}</span>
        </div>
        <h4 class="product-name">${escapeHtml(product.name)}</h4>
        <p class="product-description">${escapeHtml(product.description)}</p>
        <p class="product-price"><span>Precio de venta</span><strong>${formatCOP(product.price)}</strong></p>
        <div class="product-actions">
          <button class="button button-silver info-btn" type="button">Ver información</button>
          <button class="button button-gold add-btn" type="button" ${product.stock < 1 ? 'disabled' : ''}>
            ${product.stock < 1 ? 'Agotado' : 'Añadir al carrito'}
          </button>
        </div>
      </div>`;

    const img = article.querySelector('.product-image');
    const fallback = article.querySelector('.product-image-fallback');
    attachPhotoFallback(img, fallback, product);
    article.querySelector('.info-btn').addEventListener('click', () => openProductDetail(product));
    article.querySelector('.add-btn').addEventListener('click', () => openProductDetail(product, { focusPurchase: true }));
    return article;
  }

  function variantsForProduct(product) {
    // Mantiene el producto como variante independiente, pero si el Excel contiene
    // la misma referencia visual con varias tallas, permite elegirlas en el modal.
    return state.products.filter(p =>
      p.collection === product.collection &&
      p.category === product.category &&
      p.subcategory === product.subcategory &&
      p.descriptionBase === product.descriptionBase &&
      p.color === product.color &&
      p.gender === product.gender &&
      p.photo === product.photo &&
      p.price === product.price
    );
  }

  function openProductDetail(product, { focusPurchase = false } = {}) {
    state.selectedProduct = product;
    state.selectedVariantId = product.variantId;
    state.detailQuantity = 1;

    elements.detailTitle.textContent = product.name;
    elements.detailReference.textContent = `Referencia: ${product.referencia}`;
    elements.detailPrice.textContent = formatCOP(product.price);
    elements.detailDescription.textContent = product.descriptionBase;
    elements.detailColor.textContent = product.color || 'No especificado';
    elements.detailGender.textContent = product.gender || 'No especificado';
    attachPhotoFallback(elements.detailImage, elements.detailImageFallback, product);
    elements.detailError.hidden = true;

    const variants = variantsForProduct(product);
    elements.detailSizeSelect.innerHTML = '';
    variants.forEach(variant => {
      const option = document.createElement('option');
      option.value = variant.variantId;
      option.textContent = `${variant.size} · ${remainingStock(variant.variantId)} disponible${remainingStock(variant.variantId) === 1 ? '' : 's'}`;
      option.disabled = remainingStock(variant.variantId) < 1;
      elements.detailSizeSelect.appendChild(option);
    });

    const preferred = variants.find(v => v.variantId === product.variantId && remainingStock(v.variantId) > 0)
      || variants.find(v => remainingStock(v.variantId) > 0)
      || product;
    state.selectedVariantId = preferred.variantId;
    elements.detailSizeSelect.value = preferred.variantId;

    updateDetailVariant();
    elements.productDialog.showModal();
    if (focusPurchase) elements.detailSizeSelect.focus();
  }

  function selectedDetailVariant() {
    return state.products.find(p => p.variantId === state.selectedVariantId) || state.selectedProduct;
  }

  function updateDetailVariant() {
    const product = selectedDetailVariant();
    if (!product) return;

    state.detailQuantity = 1;
    elements.detailReference.textContent = `Referencia: ${product.referencia}`;
    elements.detailPrice.textContent = formatCOP(product.price);
    elements.detailDescription.textContent = product.descriptionBase;
    elements.detailColor.textContent = product.color || 'No especificado';
    elements.detailSizeText.textContent = product.size || 'Única';
    elements.detailGender.textContent = product.gender || 'No especificado';
    attachPhotoFallback(elements.detailImage, elements.detailImageFallback, product);
    updateDetailControls();
  }

  function updateDetailControls() {
    const product = selectedDetailVariant();
    if (!product) return;
    const remaining = remainingStock(product.variantId);
    state.detailQuantity = Math.min(Math.max(1, state.detailQuantity), Math.max(1, remaining));
    elements.detailQtyValue.textContent = String(state.detailQuantity);
    elements.detailStockNotice.textContent = remaining > 0 ? `Existencia registrada: ${remaining} unidades.` : 'Producto agotado.';
    elements.detailSubtotal.textContent = formatCOP(product.price * state.detailQuantity);
    elements.detailQtyDec.disabled = state.detailQuantity <= 1;
    elements.detailQtyInc.disabled = remaining < 1 || state.detailQuantity >= remaining;
    elements.detailAddButton.disabled = remaining < 1;
    elements.detailAddButton.textContent = remaining < 1 ? 'Agotado' : 'Agregar al carrito';
  }

  function changeDetailQuantity(delta) {
    const product = selectedDetailVariant();
    if (!product) return;
    const remaining = remainingStock(product.variantId);
    if (remaining < 1) return;
    state.detailQuantity = Math.min(Math.max(1, state.detailQuantity + delta), remaining);
    updateDetailControls();
  }

  function addSelectedProduct() {
    const product = selectedDetailVariant();
    if (!product) return;

    const quantity = Math.max(1, Math.floor(Number(state.detailQuantity) || 1));
    const remaining = remainingStock(product.variantId);
    if (quantity > remaining) {
      showError(elements.detailError, `Solo quedan ${remaining} unidades disponibles para esta variante.`);
      return;
    }

    const existing = state.cart.find(line => line.variantId === product.variantId);
    if (existing) existing.quantity += quantity;
    else state.cart.push({ variantId: product.variantId, referencia: product.referencia, size: product.size, quantity });

    persistCart();
    renderCart();
    elements.productDialog.close();
    openCart();
  }

  function findProductForLine(line) {
    if (line.variantId) {
      const exact = state.products.find(p => p.variantId === line.variantId);
      if (exact) return exact;
    }
    return state.products.find(p => p.referencia === line.referencia && p.size === line.size)
      || state.products.find(p => p.referencia === line.referencia);
  }

  function remainingStock(variantId) {
    const product = state.products.find(p => p.variantId === variantId);
    if (!product) return 0;
    const used = state.cart
      .filter(line => line.variantId === variantId)
      .reduce((sum, line) => sum + line.quantity, 0);
    return Math.max(0, product.stock - used);
  }

  function pruneInvalidCart() {
    const migrated = [];
    state.cart.forEach(line => {
      const product = findProductForLine(line);
      if (!product || line.quantity <= 0) return;
      const alreadyUsed = migrated
        .filter(item => item.variantId === product.variantId)
        .reduce((sum, item) => sum + item.quantity, 0);
      const allowed = Math.max(0, product.stock - alreadyUsed);
      const quantity = Math.min(Math.max(0, Math.floor(line.quantity)), allowed);
      if (quantity > 0) migrated.push({ variantId: product.variantId, referencia: product.referencia, size: product.size, quantity });
    });
    state.cart = migrated;
    persistCart();
  }

  function renderCart() {
    elements.cartItems.innerHTML = '';

    if (!state.cart.length) {
      elements.cartItems.innerHTML = '<div class="cart-empty">Tu carrito está vacío.</div>';
    } else {
      state.cart.forEach((line, index) => {
        const product = findProductForLine(line);
        if (!product) return;

        const node = document.createElement('div');
        node.className = 'cart-line';
        node.innerHTML = `
          <div class="cart-thumb-wrap">
            <img class="cart-thumb" alt="" />
            <div class="product-image-fallback cart-fallback" hidden></div>
          </div>
          <div>
            <h3>${escapeHtml(product.name)}</h3>
            <p>${escapeHtml(product.description)}</p>
            <p>Ref. ${escapeHtml(product.referencia)} · ${formatCOP(product.price)} c/u</p>
            <div class="cart-controls">
              <button type="button" data-dec aria-label="Disminuir">−</button>
              <strong>${line.quantity}</strong>
              <button type="button" data-inc aria-label="Aumentar">+</button>
              <button type="button" class="remove-line" data-remove>Eliminar</button>
            </div>
          </div>
          <div class="cart-line-total">${formatCOP(product.price * line.quantity)}</div>`;

        attachPhotoFallback(node.querySelector('.cart-thumb'), node.querySelector('.cart-fallback'), product);
        node.querySelector('[data-dec]').addEventListener('click', () => changeQuantity(index, -1));
        node.querySelector('[data-inc]').addEventListener('click', () => changeQuantity(index, 1));
        node.querySelector('[data-remove]').addEventListener('click', () => removeLine(index));
        elements.cartItems.appendChild(node);
      });
    }

    const count = state.cart.reduce((sum, line) => sum + line.quantity, 0);
    elements.cartCount.textContent = String(count);
    elements.cartTotal.textContent = formatCOP(cartTotal());
    elements.checkoutButton.disabled = !state.cart.length;
  }

  function changeQuantity(index, delta) {
    const line = state.cart[index];
    const product = findProductForLine(line);
    if (!product) return;
    if (delta > 0 && remainingStock(product.variantId) < 1) return;
    line.quantity += delta;
    if (line.quantity <= 0) state.cart.splice(index, 1);
    persistCart();
    renderCart();
  }

  function removeLine(index) {
    state.cart.splice(index, 1);
    persistCart();
    renderCart();
  }

  function cartTotal() {
    return state.cart.reduce((sum, line) => {
      const product = findProductForLine(line);
      return sum + (product ? product.price * line.quantity : 0);
    }, 0);
  }

  function renderSalesNumbers() {
    elements.salesWhatsapp.innerHTML = '';
    if (!state.salesNumbers.length) {
      elements.salesWhatsapp.innerHTML = '<option value="">Agrega un WhatsApp en la segunda hoja del Excel</option>';
      return;
    }

    state.salesNumbers.forEach((number, index) => {
      const option = document.createElement('option');
      option.value = number;
      option.textContent = `Ventas ${index + 1} · +${number}`;
      elements.salesWhatsapp.appendChild(option);
    });
  }

  function setZone(type) {
    const rural = type === 'rural';
    elements.urbanAddressField.hidden = rural;
    elements.ruralPlaceField.hidden = !rural;
    elements.ruralReferenceField.hidden = !rural;
    elements.urbanAddress.required = !rural;
    elements.ruralPlace.required = rural;
    elements.ruralReference.required = rural;
  }

  function openCheckout() {
    if (!state.cart.length) return;
    renderSalesNumbers();
    setZone(elements.zoneType.value);
    elements.checkoutError.hidden = true;
    elements.checkoutDialog.showModal();
  }

  function buildOrder() {
    if (!state.salesNumbers.length || !elements.salesWhatsapp.value) {
      throw new Error('No hay un WhatsApp de ventas configurado en la segunda hoja del Excel.');
    }

    const zone = elements.zoneType.value;
    const address = zone === 'urbana'
      ? `Zona urbana · ${elements.urbanAddress.value.trim()}`
      : `Zona rural · ${elements.ruralPlace.value.trim()} · ${elements.ruralReference.value.trim()}`;

    const orderId = createOrderId();
    const lines = state.cart.map((line, index) => {
      const product = findProductForLine(line);
      return [
        `${index + 1}. ${product.name}`,
        `   Descripción: ${product.description}`,
        `   Ref: ${product.referencia}`,
        `   Talla/presentación: ${product.size}`,
        `   Cantidad: ${line.quantity}`,
        `   Precio de venta: ${formatCOP(product.price)}`,
        `   Subtotal: ${formatCOP(product.price * line.quantity)}`
      ].join('\n');
    }).join('\n\n');

    const report = [
      `PEDIDO MOS · ${orderId}`,
      `Fecha: ${new Date().toLocaleString('es-CO')}`,
      '',
      'DATOS DEL COMPRADOR',
      `Nombre: ${$('customerName').value.trim()}`,
      `WhatsApp: ${$('customerPhone').value.trim()}`,
      `Entrega: ${address}`,
      'Tratamiento de datos: ACEPTADO',
      '',
      'PRODUCTOS',
      lines,
      '',
      `TOTAL: ${formatCOP(cartTotal())}`,
      '',
      'La compra continúa y se confirma por WhatsApp.'
    ].join('\n');

    return { orderId, report, salesNumber: elements.salesWhatsapp.value };
  }

  function submitCheckout() {
    try {
      const order = buildOrder();
      const url = `https://wa.me/${order.salesNumber}?text=${encodeURIComponent(order.report)}`;
      elements.reportTitle.textContent = `Pedido ${order.orderId}`;
      elements.reportText.textContent = order.report;
      elements.openWhatsappButton.href = url;
      elements.checkoutDialog.close();
      elements.reportDialog.showModal();
      window.open(url, '_blank', 'noopener');
    } catch (error) {
      showError(elements.checkoutError, error.message || 'No se pudo generar el pedido.');
    }
  }

  function createOrderId() {
    const d = new Date();
    const pad = n => String(n).padStart(2, '0');
    const date = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
    const time = `${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
    const random = Math.random().toString(36).slice(2, 6).toUpperCase();
    return `MOS-${date}-${time}-${random}`;
  }

  function openCart() {
    elements.cartDrawer.classList.add('open');
    elements.cartDrawer.setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';
  }

  function closeCart() {
    elements.cartDrawer.classList.remove('open');
    elements.cartDrawer.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
  }

  function formatCOP(value) {
    return new Intl.NumberFormat('es-CO', {
      style: 'currency',
      currency: 'COP',
      maximumFractionDigits: 0
    }).format(value || 0);
  }

  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>'"]/g, ch => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      "'": '&#39;',
      '"': '&quot;'
    }[ch]));
  }

  function showError(el, message) {
    el.textContent = message;
    el.hidden = false;
  }

  function persistCart() {
    localStorage.setItem('mos-cart-final', JSON.stringify(state.cart));
  }

  function loadCart() {
    try {
      return JSON.parse(
        localStorage.getItem('mos-cart-final')
        || localStorage.getItem('mos-cart-v3')
        || localStorage.getItem('mos-cart-v2')
        || '[]'
      );
    } catch {
      return [];
    }
  }

  elements.cartButton.addEventListener('click', openCart);
  document.querySelectorAll('[data-close-cart]').forEach(el => el.addEventListener('click', closeCart));
  document.querySelectorAll('[data-close-product]').forEach(el => el.addEventListener('click', () => elements.productDialog.close()));
  document.querySelectorAll('[data-close-checkout]').forEach(el => el.addEventListener('click', () => elements.checkoutDialog.close()));
  document.querySelectorAll('[data-close-report]').forEach(el => el.addEventListener('click', () => elements.reportDialog.close()));

  elements.detailSizeSelect.addEventListener('change', () => {
    state.selectedVariantId = elements.detailSizeSelect.value;
    elements.detailError.hidden = true;
    updateDetailVariant();
  });
  elements.detailQtyDec.addEventListener('click', () => changeDetailQuantity(-1));
  elements.detailQtyInc.addEventListener('click', () => changeDetailQuantity(1));
  elements.detailAddButton.addEventListener('click', addSelectedProduct);
  elements.checkoutButton.addEventListener('click', openCheckout);
  elements.zoneType.addEventListener('change', () => setZone(elements.zoneType.value));

  elements.checkoutForm.addEventListener('submit', event => {
    event.preventDefault();
    elements.checkoutError.hidden = true;
    if (!elements.checkoutForm.reportValidity()) return;
    submitCheckout();
  });

  elements.copyReportButton.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(elements.reportText.textContent);
      elements.copyReportButton.textContent = 'Informe copiado';
      setTimeout(() => { elements.copyReportButton.textContent = 'Copiar informe'; }, 1400);
    } catch {
      elements.copyReportButton.textContent = 'No se pudo copiar';
    }
  });

  document.addEventListener('keydown', event => {
    if (event.key === 'Escape') closeCart();
  });

  window.addEventListener('DOMContentLoaded', () => {
    loadWorkbook();
    setInterval(() => loadWorkbook({ silent: true }), POLL_MS);
  });
})();
