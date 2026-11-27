import {AbsoluteFill, Easing, Img, OffthreadVideo, Series, interpolate, staticFile, useCurrentFrame} from 'remotion';

/**
 * Demo 3 — Teams audio, fixed live.
 * Same language as Demo 1 and 2. Calm + sharp, Corporate.
 * One message: it asks, takes the screen, and switches the speaker.
 * Easing cubic-bezier(0.22, 1, 0.36, 1). Base 0.4s. Crossfade only.
 * Within a group, stagger is 2 frames. Between the three steps, a hold.
 * No overshoot. The mouse-driving wait is the cut. Mic left out.
 * endAt is absolute: the clip stays mounted for endAt - startFrom frames.
 */

export const FPS = 30;
export const WIDTH = 1920;
export const HEIGHT = 1080;
const FADE = 12;
const UNIT = 12;

const TITLE = 108;
const INFO = 240;
const ASK = 150;
// Checklist while it is still looking: checking the screen, Teams, the mouse.
const FINDING = 240;
const CONTROL = 150;
const SECOND = 150;
const VIDEO_PART = 168;
// Finished frame stays up long enough to talk through what changed.
const HOLD_PART = 312;
const RESULT = VIDEO_PART + HOLD_PART;

export const DURATION = TITLE + INFO + ASK + FINDING + CONTROL + SECOND + RESULT - FADE * 7;

const PAPER = '#fafaf8';
const INK = '#0a0a0a';
const MUTED = '#737373';
const LINE = '#d4d4d2';
const ACCENT = '#002FA7';
const FONT = '"Segoe UI", sans-serif';
const easeOut = Easing.bezier(0.22, 1, 0.36, 1);

const ASK_FROM = Math.round(7.6 * FPS);
const FINDING_FROM = Math.round(33 * FPS);
const CONTROL_FROM = Math.round(56.8 * FPS);
const SECOND_FROM = Math.round(75.2 * FPS);
const RESULT_FROM = Math.round(137.5 * FPS);

const SRC_W = 1920;
const SRC_H = 1200;
const STAGE_H = 900;
const STAGE_W = Math.round(STAGE_H * (SRC_W / SRC_H));
const STAGE_X = Math.round((WIDTH - STAGE_W) / 2);
const STAGE_Y = 44;

const rise = (frame, delay) => {
  const t = frame - delay;
  const opacity = interpolate(t, [0, UNIT], [0, 1], {
    easing: easeOut,
    extrapolateLeft: 'clamp',
    extrapolateRight: 'clamp',
  });
  const y = interpolate(t, [0, UNIT], [16, 0], {
    easing: easeOut,
    extrapolateLeft: 'clamp',
    extrapolateRight: 'clamp',
  });
  return {opacity, transform: `translateY(${y}px)`};
};

const Shot = ({frames, fadeIn, fadeOut, children}) => {
  const frame = useCurrentFrame();
  const inn = fadeIn ? interpolate(frame, [0, FADE], [0, 1], {extrapolateLeft: 'clamp', extrapolateRight: 'clamp'}) : 1;
  const out = fadeOut ? interpolate(frame, [frames - FADE, frames], [1, 0], {extrapolateLeft: 'clamp', extrapolateRight: 'clamp'}) : 1;
  return <AbsoluteFill style={{opacity: inn * out}}>{children}</AbsoluteFill>;
};

const Tag = ({n, label}) => (
  <div style={{display: 'flex', alignItems: 'center', gap: 12, fontFamily: FONT, fontSize: 15, fontWeight: 600, letterSpacing: '0.16em', textTransform: 'uppercase', color: MUTED}}>
    <span style={{border: `1px solid ${LINE}`, color: INK, padding: '4px 7px', letterSpacing: '0.08em', lineHeight: 1}}>{n}</span>
    {label}
  </div>
);

const Title = () => {
  const frame = useCurrentFrame();
  return (
    <AbsoluteFill style={{background: PAPER, padding: '88px 120px', fontFamily: FONT}}>
      <div style={{fontSize: 14, fontWeight: 600, letterSpacing: '0.18em', textTransform: 'uppercase', color: MUTED, ...rise(frame, 0)}}>
        OPS Copilot Assistant
      </div>
      <div style={{marginTop: 32, maxWidth: 820, fontSize: 60, fontWeight: 300, letterSpacing: '-0.03em', lineHeight: 1.05, color: INK, ...rise(frame, 6)}}>
        It takes control and fixes the Teams audio live.
      </div>
      <div style={{display: 'flex', gap: 56, marginTop: 64, alignItems: 'stretch'}}>
        <div style={{flex: 1, ...rise(frame, 20)}}>
          <Tag n="01" label="The situation" />
          <div style={{marginTop: 20, fontSize: 34, fontWeight: 400, letterSpacing: '-0.02em', lineHeight: 1.2, color: INK}}>
            A meeting starts soon. The headset has no sound.
          </div>
          <div style={{marginTop: 12, fontSize: 20, lineHeight: 1.4, color: MUTED}}>A device problem, not a policy question.</div>
        </div>
        <div style={{width: 1, background: LINE}} />
        <div style={{flex: 1, ...rise(frame, 22)}}>
          <Tag n="02" label="Watch for" />
          <div style={{marginTop: 20, fontSize: 34, fontWeight: 400, letterSpacing: '-0.02em', lineHeight: 1.2, color: ACCENT}}>
            It asks before it clicks. Then it solves the problem.
          </div>
          <div style={{marginTop: 12, fontSize: 20, lineHeight: 1.4, color: MUTED}}>
            It checks the audio devices, shows the step, and the Stop button stays on screen.
          </div>
        </div>
      </div>
    </AbsoluteFill>
  );
};

const Step = ({at, n, title, body, accent}) => {
  const frame = useCurrentFrame();
  return (
    <div style={{flex: 1, fontFamily: FONT}}>
      <div
        style={{
          width: 44,
          height: 44,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          border: `1px solid ${accent ? ACCENT : LINE}`,
          background: accent ? ACCENT : 'transparent',
          color: accent ? '#fff' : INK,
          fontSize: 16,
          fontWeight: 600,
          letterSpacing: '0.06em',
          ...rise(frame, at),
        }}
      >
        {n}
      </div>
      <div style={{marginTop: 18, fontSize: 32, fontWeight: 400, letterSpacing: '-0.02em', color: accent ? ACCENT : INK, ...rise(frame, at + 2)}}>{title}</div>
      <div style={{marginTop: 10, maxWidth: 340, fontSize: 20, lineHeight: 1.4, color: MUTED, ...rise(frame, at + 4)}}>{body}</div>
    </div>
  );
};

const Connector = ({at}) => {
  const frame = useCurrentFrame();
  const w = interpolate(frame, [at, at + UNIT], [0, 1], {
    easing: easeOut,
    extrapolateLeft: 'clamp',
    extrapolateRight: 'clamp',
  });
  return (
    <div style={{width: 72, marginTop: 21, flex: '0 0 72px'}}>
      <div style={{height: 2, background: ACCENT, transform: `scaleX(${w})`, transformOrigin: 'left center'}} />
    </div>
  );
};

const How = () => {
  const frame = useCurrentFrame();
  return (
    <AbsoluteFill style={{background: PAPER, padding: '88px 120px', fontFamily: FONT}}>
      <div style={{fontSize: 14, fontWeight: 600, letterSpacing: '0.18em', textTransform: 'uppercase', color: MUTED, ...rise(frame, 0)}}>
        How the sound comes back
      </div>
      <div style={{marginTop: 28, maxWidth: 760, fontSize: 52, fontWeight: 300, letterSpacing: '-0.03em', lineHeight: 1.08, color: INK, ...rise(frame, 6)}}>
        Ask. Then the screen. Then the speaker.
      </div>
      <div style={{display: 'flex', alignItems: 'flex-start', marginTop: 72}}>
        <Step at={36} n="01" title="Ask" body="It asks before it clicks." />
        <Connector at={68} />
        <Step at={96} n="02" title="Take control" body="It opens the Teams audio settings. Stop stays on screen." />
        <Connector at={128} />
        <Step at={156} n="03" title="Switch" body="It changes the speaker after another yes." accent />
      </div>
    </AbsoluteFill>
  );
};

const Screen = ({startFrom, endAt, style}) => (
  <OffthreadVideo src={staticFile('demo3.mp4')} startFrom={startFrom} endAt={endAt} muted style={style} />
);

const Stage = ({startFrom, endAt, hold, enter, children}) => {
  const frame = useCurrentFrame();
  const scale = enter
    ? interpolate(frame, [0, UNIT], [0.96, 1], {easing: easeOut, extrapolateLeft: 'clamp', extrapolateRight: 'clamp'})
    : 1;
  const mediaStyle = {position: 'absolute', left: 0, top: 0, width: STAGE_W, height: STAGE_H, objectFit: 'fill'};
  const bandTop = Math.round(STAGE_H * 0.1);
  const bandH = Math.round(STAGE_H * 0.022);
  const bandLeft = Math.round(STAGE_W * 0.02);
  const bandW = Math.round(STAGE_W * 0.41);
  const media = (style) =>
    hold ? <Img src={staticFile('demo3-hold.png')} style={style} /> : <Screen startFrom={startFrom} endAt={endAt} style={style} />;
  return (
    <AbsoluteFill style={{background: PAPER}}>
      <div style={{position: 'absolute', top: 16, left: 88, fontFamily: FONT, fontSize: 13, fontWeight: 600, letterSpacing: '0.18em', textTransform: 'uppercase', color: MUTED}}>
        OPS Copilot Assistant
      </div>
      <div style={{position: 'absolute', top: 16, right: 88, fontFamily: FONT, fontSize: 13, fontWeight: 600, letterSpacing: '0.18em', textTransform: 'uppercase', color: MUTED}}>
        Demo 3
      </div>
      <div
        style={{
          position: 'absolute',
          left: STAGE_X,
          top: STAGE_Y,
          width: STAGE_W,
          height: STAGE_H,
          transform: `scale(${scale})`,
          transformOrigin: '50% 50%',
          borderRadius: 12,
          overflow: 'hidden',
          background: '#fff',
          boxShadow: '0 18px 50px rgba(10,10,10,0.10)',
        }}
      >
        {media(mediaStyle)}
        <div style={{position: 'absolute', left: bandLeft, top: bandTop, width: bandW, height: bandH, overflow: 'hidden'}}>
          {media({...mediaStyle, left: -bandLeft, top: -bandTop, filter: 'blur(5px)'})}
          <div style={{position: 'absolute', inset: 0, background: 'rgba(244,244,242,0.78)'}} />
        </div>
        {children}
      </div>
    </AbsoluteFill>
  );
};

const Caption = ({n, label, line, color}) => (
  <div style={{position: 'absolute', left: STAGE_X, right: WIDTH - STAGE_X - STAGE_W, bottom: 28, display: 'flex', alignItems: 'baseline', gap: 22, fontFamily: FONT}}>
    <Tag n={n} label={label} />
    <div style={{fontSize: 26, fontWeight: 400, letterSpacing: '-0.02em', color, lineHeight: 1.2}}>{line}</div>
  </div>
);

const Mark = () => {
  const frame = useCurrentFrame();
  const scaleY = interpolate(frame, [24, 36], [0, 1], {easing: easeOut, extrapolateLeft: 'clamp', extrapolateRight: 'clamp'});
  return (
    <div
      style={{
        position: 'absolute',
        left: STAGE_W * 0.685,
        top: STAGE_H * 0.16,
        width: 3,
        height: STAGE_H * 0.2,
        background: ACCENT,
        transform: `scaleY(${scaleY})`,
        transformOrigin: 'top center',
      }}
    />
  );
};

const ResultBody = () => {
  const frame = useCurrentFrame();
  return (
    <Shot frames={RESULT} fadeIn fadeOut={false}>
      <Stage startFrom={RESULT_FROM} endAt={RESULT_FROM + VIDEO_PART} hold={frame >= VIDEO_PART}>
        <Mark />
      </Stage>
      <Caption n="03" label="Fixed" line="The meeting audio leaves the monitor." color={ACCENT} />
    </Shot>
  );
};

export const Demo3 = () => (
  <AbsoluteFill style={{background: PAPER}}>
    <Series>
      <Series.Sequence durationInFrames={TITLE}>
        <Shot frames={TITLE} fadeIn={false} fadeOut>
          <Title />
        </Shot>
      </Series.Sequence>
      <Series.Sequence durationInFrames={INFO} offset={-FADE}>
        <Shot frames={INFO} fadeIn fadeOut>
          <How />
        </Shot>
      </Series.Sequence>
      <Series.Sequence durationInFrames={ASK} offset={-FADE}>
        <Shot frames={ASK} fadeIn fadeOut>
          <Stage startFrom={ASK_FROM} endAt={ASK_FROM + ASK} enter />
          <Caption n="01" label="Ask" line="It asks before it clicks." color={INK} />
        </Shot>
      </Series.Sequence>
      <Series.Sequence durationInFrames={FINDING} offset={-FADE}>
        <Shot frames={FINDING} fadeIn fadeOut>
          <Stage startFrom={FINDING_FROM} endAt={FINDING_FROM + FINDING} />
          <Caption n="01" label="Looking" line="It shows each step while it looks." color={INK} />
        </Shot>
      </Series.Sequence>
      <Series.Sequence durationInFrames={CONTROL} offset={-FADE}>
        <Shot frames={CONTROL} fadeIn fadeOut>
          <Stage startFrom={CONTROL_FROM} endAt={CONTROL_FROM + CONTROL} />
          <Caption n="02" label="Take control" line="It opens the Teams audio settings." color={ACCENT} />
        </Shot>
      </Series.Sequence>
      <Series.Sequence durationInFrames={SECOND} offset={-FADE}>
        <Shot frames={SECOND} fadeIn fadeOut>
          <Stage startFrom={SECOND_FROM} endAt={SECOND_FROM + SECOND} />
          <Caption n="02" label="Ask again" line="It asks again, before it switches." color={ACCENT} />
        </Shot>
      </Series.Sequence>
      <Series.Sequence durationInFrames={RESULT} offset={-FADE}>
        <ResultBody />
      </Series.Sequence>
    </Series>
  </AbsoluteFill>
);
