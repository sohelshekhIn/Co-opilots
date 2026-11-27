import {Composition} from 'remotion';
import {Demo1, DURATION, FPS, HEIGHT, WIDTH} from './Demo1';
import {Demo2, DURATION as DURATION_2, FPS as FPS_2, HEIGHT as HEIGHT_2, WIDTH as WIDTH_2} from './Demo2';
import {Demo3, DURATION as DURATION_3, FPS as FPS_3, HEIGHT as HEIGHT_3, WIDTH as WIDTH_3} from './Demo3';

export const Root = () => {
  return (
    <>
      <Composition
        id="Demo1"
        component={Demo1}
        durationInFrames={DURATION}
        fps={FPS}
        width={WIDTH}
        height={HEIGHT}
      />
      <Composition
        id="Demo2"
        component={Demo2}
        durationInFrames={DURATION_2}
        fps={FPS_2}
        width={WIDTH_2}
        height={HEIGHT_2}
      />
      <Composition
        id="Demo3"
        component={Demo3}
        durationInFrames={DURATION_3}
        fps={FPS_3}
        width={WIDTH_3}
        height={HEIGHT_3}
      />
    </>
  );
};
