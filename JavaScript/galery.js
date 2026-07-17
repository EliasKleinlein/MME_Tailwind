const DATA_FILE = "Daten/fahrzeuge.xlsx";
const GALLERY_ID = "vehicle-gallery";

document.addEventListener("DOMContentLoaded", loadVehicles);

async function loadVehicles() {
  const gallery = document.getElementById(GALLERY_ID);

  if (!gallery) {
    console.error(`Container #${GALLERY_ID} wurde nicht gefunden.`);
    return;
  }

  try {
    const response = await fetch(DATA_FILE);

    if (!response.ok) {
      throw new Error(
        `Excel-Datei konnte nicht geladen werden: ${response.status} ${response.statusText}`,
      );
    }

    const workbookBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(workbookBuffer, { type: "array" });
    const worksheet = workbook.Sheets["Fahrzeuge"];

    if (!worksheet) {
      throw new Error('Das Tabellenblatt "Fahrzeuge" fehlt.');
    }

    const vehicles = XLSX.utils
      .sheet_to_json(worksheet, {
        defval: "",
        raw: false,
      })
      .filter((vehicle) => normalizeText(vehicle.aktiv) === "JA")
      .sort((a, b) => Number(a.position) - Number(b.position));

    gallery.innerHTML = vehicles
      .map((vehicle, index) => createVehicleCard(vehicle, index))
      .join("");
  } catch (error) {
    console.error(error);

    gallery.innerHTML = `
      <div class="rounded-2xl border border-red-500/40 bg-red-950/40 p-6 text-center text-red-200">
        Die Fahrzeugdaten konnten nicht geladen werden.
      </div>
    `;
  }
}

function createVehicleCard(vehicle, index) {
  const displayNumber = String(index + 1).padStart(2, "0");
  const zoomId = `vehicle-zoom-${index + 1}`;
  const isImageLeft = index % 2 === 0;
  const textClass = String(vehicle.textklasse).trim();

  const accent = [
    toRgbNumber(vehicle.akzent_r),
    toRgbNumber(vehicle.akzent_g),
    toRgbNumber(vehicle.akzent_b),
  ].join(" ");

  const image1 = buildImagePath(vehicle.bildname, 1, vehicle.bild1_typ);

  const image2 = buildImagePath(vehicle.bildname, 2, vehicle.bild2_typ);

  const media = `
    <label
      class="vehicle-media ${
        isImageLeft ? "" : "order-1 lg:order-2"
      } rounded-[36px]"
      for="${escapeHtml(zoomId)}"
    >
      <img
        alt="${escapeHtml(vehicle.hersteller)} ${escapeHtml(vehicle.modell)} – Ansicht 1"
        class="vehicle-image vehicle-image-primary"
        src="${escapeHtml(image1)}"
      />

      <img
        alt="${escapeHtml(vehicle.hersteller)} ${escapeHtml(vehicle.modell)} – Ansicht 2"
        class="vehicle-image vehicle-image-secondary"
        src="${escapeHtml(image2)}"
      />

      <span
        class="vehicle-number absolute bottom-5 ${
          isImageLeft ? "left-7" : "right-7"
        } z-10 title text-7xl"
      >
        ${displayNumber}
      </span>
    </label>
  `;

  const info = `
    <div
      class="vehicle-info ${
        isImageLeft
          ? "rounded-[30px] p-7 lg:-ml-20 lg:p-9"
          : "order-2 rounded-[30px] p-7 lg:order-1 lg:-mr-20 lg:p-9"
      }"
    >
      <p
        class="text-xs font-semibold uppercase tracking-[0.4em] ${escapeHtml(textClass)}"
      >
        ${escapeHtml(vehicle.hersteller)}
      </p>

      <h2 class="title mt-2 text-5xl tracking-wider">
        ${escapeHtml(vehicle.modell)}
      </h2>

      <p class="mt-4 leading-7 text-zinc-300">
        ${escapeHtml(vehicle.beschreibung)}
      </p>

      <div class="vehicle-details">
        <div class="my-6 h-px bg-white/10"></div>

        <div class="grid grid-cols-3 gap-3 text-center">
          ${createDataItem(vehicle.wert1, vehicle.label1, textClass)}
          ${createDataItem(vehicle.wert2, vehicle.label2, textClass)}
          ${createDataItem(vehicle.wert3, vehicle.label3, textClass)}
        </div>

        <a
          class="vehicle-calendar mt-7 flex items-center justify-between rounded-2xl border px-5 py-4 transition duration-500"
          href="${escapeHtml(vehicle.kontaktlink || "kontact.html#termin")}"
        >
          <span>
            <strong class="block">Verfügbarkeit prüfen</strong>
            <small class="text-zinc-400">
              Kalender und Mietanfrage
            </small>
          </span>

          <svg
            aria-hidden="true"
            class="h-7 w-7"
            fill="none"
            stroke="currentColor"
            stroke-width="1.7"
            viewBox="0 0 24 24"
          >
            <rect height="16" rx="2" width="18" x="3" y="5"></rect>
            <path d="M16 3v4M8 3v4M3 10h18"></path>
            <path d="m9 16 2 2 4-4"></path>
          </svg>
        </a>
      </div>
    </div>
  `;

  return `
    <article
      class="vehicle grid items-center gap-6 lg:grid-cols-2"
      style="--accent: ${accent}"
    >
      <input
        class="zoom-toggle"
        id="${escapeHtml(zoomId)}"
        type="checkbox"
      />

      ${isImageLeft ? `${media}${info}` : `${info}${media}`}
    </article>
  `;
}

function createDataItem(value, label, textClass) {
  return `
    <div>
      <strong class="text-xl ${escapeHtml(textClass)}">
        ${escapeHtml(value)}
      </strong>

      <p class="text-xs text-zinc-500">
        ${escapeHtml(label)}
      </p>
    </div>
  `;
}

function buildImagePath(imageName, number, extension) {
  const cleanName = String(imageName).trim();

  const cleanExtension = String(extension)
    .trim()
    .replace(/^\./, "")
    .toLowerCase();

  const separator = cleanName.endsWith("_") ? "" : " ";

  return `Galeriebilder/${cleanName}${separator}${number}.${cleanExtension}`;
}

function normalizeText(value) {
  return String(value).trim().toUpperCase();
}

function toRgbNumber(value) {
  const number = Number(value);

  if (!Number.isFinite(number)) {
    return 0;
  }

  return Math.min(255, Math.max(0, Math.round(number)));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
